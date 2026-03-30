/**
 * FullscreenTimeline Component
 * Zoomable fullscreen timeline for the camera view.
 */

import { h } from 'preact';
import { useState, useEffect, useRef, useCallback, useMemo } from 'preact/hooks';

function formatDateISO(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatDayLabel(date, offsetDays = 0) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + offsetDays);
  return copy.toLocaleDateString(undefined, {
    day: 'numeric',
    month: 'long',
    year: 'numeric'
  });
}

function formatHourLabel(timestamp, intervalSeconds) {
  const date = new Date(timestamp * 1000);
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = date.getSeconds();
  const millis = date.getMilliseconds();

  if (intervalSeconds >= 3600) {
    const suffix = hours >= 12 ? 'PM' : 'AM';
    const hour12 = hours % 12 || 12;
    return `${hour12} ${suffix}`;
  }

  if (intervalSeconds >= 60) {
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
  }

  if (intervalSeconds >= 1) {
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
  }

  return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}.${String(millis).padStart(3, '0')}`;
}

function formatPreciseTime(timestamp) {
  const date = new Date(timestamp * 1000);
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  const millis = String(date.getMilliseconds()).padStart(3, '0');
  return `${hours}:${minutes}:${seconds}.${millis}`;
}

function getTickInterval(windowSeconds) {
  if (windowSeconds >= 24 * 3600) return 3600;
  if (windowSeconds >= 6 * 3600) return 1800;
  if (windowSeconds >= 3600) return 900;
  if (windowSeconds >= 10 * 60) return 60;
  if (windowSeconds >= 2 * 60) return 10;
  if (windowSeconds >= 20) return 1;
  if (windowSeconds >= 2) return 0.1;
  return 0.05;
}

export function FullscreenTimeline({ streamName, videoRef }) {
  const ZOOM_WINDOWS = [86400, 21600, 3600, 300, 60, 10, 1];

  const [segments, setSegments] = useState([]);
  const [detections, setDetections] = useState([]);
  const [mergedSegments, setMergedSegments] = useState([]);
  const [mergedDetections, setMergedDetections] = useState([]);
  const [visible, setVisible] = useState(true);
  const [isPlayingRecording, setIsPlayingRecording] = useState(false);
  const [playbackTime, setPlaybackTime] = useState(null);
  const [activeSegment, setActiveSegment] = useState(null);
  const [recordingPaused, setRecordingPaused] = useState(false);
  const [zoomIndex, setZoomIndex] = useState(1);
  const [cursorTimestamp, setCursorTimestamp] = useState(Date.now() / 1000);
  const [isDraggingCursor, setIsDraggingCursor] = useState(false);
  const [skipIndicator, setSkipIndicator] = useState({ visible: false, text: '' });

  const containerRef = useRef(null);
  const hideTimerRef = useRef(null);
  const nowIntervalRef = useRef(null);
  const recordingVideoRef = useRef(null);
  const dragTimestampRef = useRef(cursorTimestamp);
  const zoomAccumulatorRef = useRef(0);

  const HIDE_DELAY_MS = 4000;
  const currentDate = useMemo(() => new Date(), []);
  const dateStr = useMemo(() => formatDateISO(currentDate), [currentDate]);
  const [year, month, day] = dateStr.split('-').map(Number);
  const dayStartTimestamp = useMemo(
    () => new Date(year, month - 1, day, 0, 0, 0, 0).getTime() / 1000,
    [year, month, day]
  );
  const dayEndTimestamp = dayStartTimestamp + 24 * 3600;

  useEffect(() => {
    if (!streamName) return;

    const startDate = new Date(year, month - 1, day, 0, 0, 0, 0);
    const endDate = new Date(year, month - 1, day, 23, 59, 59, 999);
    const url = `/api/timeline/segments?stream=${encodeURIComponent(streamName)}&start=${encodeURIComponent(startDate.toISOString())}&end=${encodeURIComponent(endDate.toISOString())}`;

    fetch(url)
      .then(res => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
      })
      .then(data => {
        setSegments(data.segments || []);
        setDetections(data.detections || []);
      })
      .catch(err => {
        console.error('FullscreenTimeline: Error fetching segments:', err);
        setSegments([]);
        setDetections([]);
      });
  }, [streamName, year, month, day]);

  useEffect(() => {
    if (!segments.length) {
      setMergedSegments([]);
      return;
    }

    const sorted = [...segments].sort((a, b) => a.start_timestamp - b.start_timestamp);
    const merged = [];
    let current = { ...sorted[0], originalSegments: [sorted[0]] };

    for (let i = 1; i < sorted.length; i++) {
      const seg = sorted[i];
      const gap = seg.start_timestamp - current.end_timestamp;
      if (gap <= 2) {
        current.end_timestamp = Math.max(current.end_timestamp, seg.end_timestamp);
        current.originalSegments.push(seg);
        current.has_detection = current.has_detection || seg.has_detection;
      } else {
        merged.push(current);
        current = { ...seg, originalSegments: [seg] };
      }
    }

    merged.push(current);
    setMergedSegments(merged);
  }, [segments]);

  useEffect(() => {
    if (!detections.length) {
      setMergedDetections([]);
      return;
    }

    const sorted = [...detections].sort((a, b) => a.start_timestamp - b.start_timestamp);
    const merged = [];
    let current = { ...sorted[0] };

    for (let i = 1; i < sorted.length; i++) {
      const det = sorted[i];
      const gap = det.start_timestamp - current.end_timestamp;
      if (gap <= 2) {
        current.end_timestamp = Math.max(current.end_timestamp, det.end_timestamp);
      } else {
        merged.push(current);
        current = { ...det };
      }
    }

    merged.push(current);
    setMergedDetections(merged);
  }, [detections]);

  useEffect(() => {
    const update = () => {
      if (!isPlayingRecording && !isDraggingCursor) {
        setCursorTimestamp(Date.now() / 1000);
      }
    };
    update();
    nowIntervalRef.current = setInterval(update, 200);
    return () => clearInterval(nowIntervalRef.current);
  }, [isPlayingRecording, isDraggingCursor]);

  useEffect(() => {
    if (!isDraggingCursor) {
      const nextCursor = isPlayingRecording && playbackTime ? playbackTime : Date.now() / 1000;
      setCursorTimestamp(prev => {
        if (Math.abs(prev - nextCursor) < 0.02) return prev;
        return nextCursor;
      });
    }
  }, [isPlayingRecording, playbackTime, isDraggingCursor]);

  const resetHideTimer = useCallback(() => {
    setVisible(true);
    if (hideTimerRef.current) clearTimeout(hideTimerRef.current);
    hideTimerRef.current = setTimeout(() => setVisible(false), HIDE_DELAY_MS);
  }, []);

  useEffect(() => {
    const handleMouseMove = () => resetHideTimer();
    document.addEventListener('mousemove', handleMouseMove);
    resetHideTimer();

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      if (hideTimerRef.current) clearTimeout(hideTimerRef.current);
    };
  }, [resetHideTimer]);

  const playRecording = useCallback((segment, seekTime = 0) => {
    setIsPlayingRecording(true);
    setActiveSegment(segment);
    setPlaybackTime(segment.start_timestamp + seekTime);
    setRecordingPaused(false);
    setCursorTimestamp(segment.start_timestamp + seekTime);

    if (videoRef?.current) {
      videoRef.current.style.display = 'none';
    }

    const recVideo = recordingVideoRef.current;
    if (!recVideo) return;

    recVideo.style.display = 'block';
    recVideo.src = `/api/recordings/play/${segment.id}?t=${Date.now()}`;

    const onMetadata = () => {
      recVideo.currentTime = Math.min(seekTime, recVideo.duration || seekTime);
      recVideo.play().catch(err => console.error('FullscreenTimeline: playback error:', err));
      recVideo.removeEventListener('loadedmetadata', onMetadata);
    };

    recVideo.addEventListener('loadedmetadata', onMetadata);
    recVideo.load();
  }, [videoRef]);

  const returnToLive = useCallback(() => {
    setIsPlayingRecording(false);
    setActiveSegment(null);
    setPlaybackTime(null);
    setRecordingPaused(false);
    setCursorTimestamp(Date.now() / 1000);

    const recVideo = recordingVideoRef.current;
    if (recVideo) {
      recVideo.pause();
      recVideo.removeAttribute('src');
      recVideo.load();
      recVideo.style.display = 'none';
    }

    if (videoRef?.current) {
      videoRef.current.style.display = 'block';
    }
  }, [videoRef]);

  const seekToTimestamp = useCallback((timestamp) => {
    let clamped = clamp(timestamp, dayStartTimestamp, dayEndTimestamp);
    setCursorTimestamp(clamped);

    if (isPlayingRecording && activeSegment && clamped >= activeSegment.start_timestamp && clamped <= activeSegment.end_timestamp) {
      const recVideo = recordingVideoRef.current;
      if (recVideo) {
        recVideo.currentTime = Math.max(0, clamped - activeSegment.start_timestamp);
        setPlaybackTime(clamped);
      }
      return;
    }

    let foundSeg = segments.find(seg => clamped >= seg.start_timestamp && clamped <= seg.end_timestamp) || null;

    if (!foundSeg && segments.length > 0) {
      let closest = null;
      let minDist = Infinity;
      for (const seg of segments) {
        const distance = Math.min(
          Math.abs(clamped - seg.start_timestamp),
          Math.abs(clamped - seg.end_timestamp)
        );
        if (distance < minDist) {
          minDist = distance;
          closest = seg;
        }
      }
      foundSeg = closest;

      // Snap the cursor exactly to the edges of the closest segment
      if (clamped < foundSeg.start_timestamp) {
        clamped = foundSeg.start_timestamp;
      } else if (clamped > foundSeg.end_timestamp) {
        // If they click far past the very last recorded segment, return to Live
        if (foundSeg === segments[segments.length - 1] && clamped > foundSeg.end_timestamp + 2) {
            returnToLive();
            return;
        }
        clamped = Math.max(foundSeg.start_timestamp, foundSeg.end_timestamp - 0.5);
      }
      setCursorTimestamp(clamped);
    }

    if (!foundSeg) {
      return;
    }

    const seekTime = clamp(clamped - foundSeg.start_timestamp, 0, foundSeg.end_timestamp - foundSeg.start_timestamp);
    playRecording(foundSeg, seekTime);
  }, [segments, activeSegment, isPlayingRecording, playRecording, dayStartTimestamp, dayEndTimestamp]);

  const zoomWindowSeconds = ZOOM_WINDOWS[zoomIndex];
  const visibleWindowSeconds = Math.min(zoomWindowSeconds, dayEndTimestamp - dayStartTimestamp);
  const centerTimestamp = clamp(cursorTimestamp, dayStartTimestamp, dayEndTimestamp);
  const maxStart = Math.max(dayStartTimestamp, dayEndTimestamp - visibleWindowSeconds);
  const visibleStartTimestamp = visibleWindowSeconds >= (dayEndTimestamp - dayStartTimestamp)
    ? dayStartTimestamp
    : clamp(centerTimestamp - visibleWindowSeconds / 2, dayStartTimestamp, maxStart);
  const visibleEndTimestamp = visibleStartTimestamp + visibleWindowSeconds;
  const tickInterval = getTickInterval(visibleWindowSeconds);

  const timestampToPercent = useCallback((timestamp) => {
    return ((timestamp - visibleStartTimestamp) / (visibleEndTimestamp - visibleStartTimestamp)) * 100;
  }, [visibleStartTimestamp, visibleEndTimestamp]);

  const eventToTimestamp = useCallback((event) => {
    const bar = containerRef.current?.querySelector('.fs-timeline-bar');
    if (!bar) return null;

    const rect = bar.getBoundingClientRect();
    const pct = clamp((event.clientX - rect.left) / rect.width, 0, 1);
    return visibleStartTimestamp + pct * (visibleEndTimestamp - visibleStartTimestamp);
  }, [visibleStartTimestamp, visibleEndTimestamp]);

  const handleBarPointerDown = useCallback((event) => {
    const nextTimestamp = eventToTimestamp(event);
    if (nextTimestamp === null) return;

    dragTimestampRef.current = nextTimestamp;
    setIsDraggingCursor(true);
    setCursorTimestamp(nextTimestamp);

    const handleMove = (moveEvent) => {
      const moveTimestamp = eventToTimestamp(moveEvent);
      if (moveTimestamp === null) return;
      dragTimestampRef.current = moveTimestamp;
      setCursorTimestamp(moveTimestamp);
    };

    const handleUp = () => {
      setIsDraggingCursor(false);
      seekToTimestamp(dragTimestampRef.current);
      document.removeEventListener('mousemove', handleMove);
      document.removeEventListener('mouseup', handleUp);
    };

    document.addEventListener('mousemove', handleMove);
    document.addEventListener('mouseup', handleUp);
  }, [eventToTimestamp, seekToTimestamp]);

  const handleRecordingTimeUpdate = useCallback(() => {
    const recVideo = recordingVideoRef.current;
    if (!recVideo || !activeSegment) return;

    const precisePlaybackTime = activeSegment.start_timestamp + recVideo.currentTime;
    setPlaybackTime(precisePlaybackTime);
    setRecordingPaused(recVideo.paused);

    if (!isDraggingCursor) {
      setCursorTimestamp(precisePlaybackTime);
    }
  }, [activeSegment, isDraggingCursor]);

  const handleRecordingEnded = useCallback(() => {
    if (!activeSegment) {
      returnToLive();
      return;
    }

    const currentIdx = segments.findIndex(s => s.id === activeSegment.id);
    if (currentIdx >= 0 && currentIdx < segments.length - 1) {
      playRecording(segments[currentIdx + 1], 0);
    } else {
      returnToLive();
    }
  }, [activeSegment, segments, playRecording, returnToLive]);

  const toggleRecordingPlayback = useCallback(() => {
    const recVideo = recordingVideoRef.current;
    if (!recVideo || !isPlayingRecording) return;

    if (recVideo.paused) {
      recVideo.play().catch(err => console.error('FullscreenTimeline: resume playback error:', err));
      setRecordingPaused(false);
    } else {
      recVideo.pause();
      setRecordingPaused(true);
    }
  }, [isPlayingRecording]);

  const skipTimeoutRef = useRef(null);

  const showSkipIndicator = useCallback((text) => {
    setSkipIndicator({ visible: true, text });
    if (skipTimeoutRef.current) clearTimeout(skipTimeoutRef.current);
    skipTimeoutRef.current = setTimeout(() => {
      setSkipIndicator({ visible: false, text: '' });
    }, 600);
  }, []);

  const nudgeCursor = useCallback((deltaSeconds) => {
    // Determine target directly and intelligently jump gaps
    let target = cursorTimestamp + deltaSeconds;
    
    // If we're live and trying to jump forward, do nothing
    if (!isPlayingRecording && deltaSeconds > 0) return;

    let targetSeg = segments.find(s => target >= s.start_timestamp && target <= s.end_timestamp);
    
    if (!targetSeg) {
      if (deltaSeconds < 0) {
         const segsBefore = segments.filter(s => s.end_timestamp < target);
         if (segsBefore.length > 0) {
            const lastBefore = segsBefore[segsBefore.length - 1];
            if (!isPlayingRecording) {
               // Coming from LIVE: jump backward into the last recording
               target = Math.max(lastBefore.start_timestamp, lastBefore.end_timestamp + deltaSeconds);
            } else {
               // Skipping back landed in a gap, snap to latest end
               target = Math.max(lastBefore.start_timestamp, lastBefore.end_timestamp - 0.5);
            }
         }
      } else if (deltaSeconds > 0) {
         const segsAfter = segments.filter(s => s.start_timestamp > target);
         if (segsAfter.length > 0) {
            target = segsAfter[0].start_timestamp;
         } else {
            returnToLive();
            return;
         }
      }
    }
    
    showSkipIndicator(deltaSeconds > 0 ? `+${deltaSeconds}s` : `${deltaSeconds}s`);
    seekToTimestamp(target);
  }, [cursorTimestamp, seekToTimestamp, segments, isPlayingRecording, returnToLive, showSkipIndicator]);

  useEffect(() => {
    const handleKeyDown = (e) => {
      // Don't intercept if user is typing in an input field
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

      switch (e.key) {
        case 'ArrowLeft':
          e.preventDefault();
          nudgeCursor(-10);
          resetHideTimer();
          break;
        case 'ArrowRight':
          e.preventDefault();
          nudgeCursor(10);
          resetHideTimer();
          break;
        case ' ': // Spacebar
          e.preventDefault();
          toggleRecordingPlayback();
          resetHideTimer();
          break;
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [nudgeCursor, toggleRecordingPlayback, resetHideTimer]);

  const zoomIn = useCallback(() => {
    setZoomIndex(prev => Math.min(prev + 1, ZOOM_WINDOWS.length - 1));
  }, []);

  const zoomOut = useCallback(() => {
    setZoomIndex(prev => Math.max(prev - 1, 0));
  }, []);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;

    const handleWheel = (e) => {
      // If purely vertical scrolling or pinch-to-zoom (ctrlKey is true on trackpad pinch)
      if (Math.abs(e.deltaY) > Math.abs(e.deltaX) || e.ctrlKey) {
        e.preventDefault(); // Stop page scroll/zoom
        zoomAccumulatorRef.current += e.deltaY;
        
        // Threshold check to avoid jumping 50 levels on smooth scroll
        if (zoomAccumulatorRef.current > 60) {
          zoomOut();
          zoomAccumulatorRef.current = 0;
        } else if (zoomAccumulatorRef.current < -60) {
          zoomIn();
          zoomAccumulatorRef.current = 0;
        }
      }
    };

    // Passive false is required to call preventDefault on wheel events
    el.addEventListener('wheel', handleWheel, { passive: false });
    return () => el.removeEventListener('wheel', handleWheel);
  }, [zoomIn, zoomOut]);

  const renderHourMarkers = () => {
    const markers = [];
    const firstTick = Math.ceil(visibleStartTimestamp / tickInterval) * tickInterval;

    for (let tick = firstTick; tick <= visibleEndTimestamp + tickInterval / 2; tick += tickInterval) {
      const pct = timestampToPercent(tick);
      const currentHour = new Date(tick * 1000).getHours();
      const isBoundaryHour = currentHour === 0 && new Date(tick * 1000).getMinutes() === 0;
      const isFocused = Math.abs(tick - cursorTimestamp) <= tickInterval / 2;

      markers.push(
        <div
          key={`tick-${tick.toFixed(3)}`}
          className={`fs-timeline-tick ${isBoundaryHour ? 'is-boundary' : ''} ${isFocused ? 'is-focused' : ''}`}
          style={{ left: `${pct}%` }}
        />
      );

      markers.push(
        <div
          key={`label-${tick.toFixed(3)}`}
          className={`fs-timeline-label ${isFocused ? 'is-focused' : ''}`}
          style={{ left: `${pct}%` }}
        >
          {formatHourLabel(tick, tickInterval)}
        </div>
      );
    }

    return markers;
  };

  const renderSegments = () => mergedSegments
    .filter(seg => seg.end_timestamp >= visibleStartTimestamp && seg.start_timestamp <= visibleEndTimestamp)
    .map((seg, i) => {
      const leftPct = timestampToPercent(Math.max(seg.start_timestamp, visibleStartTimestamp));
      const rightPct = timestampToPercent(Math.min(seg.end_timestamp, visibleEndTimestamp));

      return (
        <div
          key={`seg-${i}`}
          className="fs-timeline-segment"
          style={{
            left: `${leftPct}%`,
            width: `${Math.max(rightPct - leftPct, 0.3)}%`
          }}
          title={`${new Date(seg.start_timestamp * 1000).toLocaleTimeString()} - ${new Date(seg.end_timestamp * 1000).toLocaleTimeString()}`}
        />
      );
    });

  const renderDetections = () => mergedDetections
    .filter(det => det.end_timestamp >= visibleStartTimestamp && det.start_timestamp <= visibleEndTimestamp)
    .map((det, i) => {
      const leftPct = timestampToPercent(Math.max(det.start_timestamp, visibleStartTimestamp));
      const rightPct = timestampToPercent(Math.min(det.end_timestamp, visibleEndTimestamp));

      return (
        <div
          key={`det-${i}`}
          className="fs-timeline-detection"
          style={{
            left: `${leftPct}%`,
            width: `${Math.max(rightPct - leftPct, 0.3)}%`
          }}
          title={`Motion: ${new Date(det.start_timestamp * 1000).toLocaleTimeString()} - ${new Date(det.end_timestamp * 1000).toLocaleTimeString()}`}
        />
      );
    });

  const renderPositionIndicator = () => {
    const pct = timestampToPercent(cursorTimestamp);
    if (pct < 0 || pct > 100) return null;

    return (
      <div className="fs-timeline-indicator" style={{ left: `${pct}%` }}>
        <div className="fs-timeline-indicator-handle"></div>
        <div className="fs-timeline-indicator-line"></div>
      </div>
    );
  };

  return (
    <>
      {/* Skip Indicator Overlay */}
      <div 
        style={{
          position: 'fixed',
          top: '50%',
          left: '50%',
          transform: 'translate(-50%, -50%)',
          backgroundColor: 'rgba(0,0,0,0.7)',
          backdropFilter: 'blur(4px)',
          color: 'white',
          padding: '20px 40px',
          borderRadius: '24px',
          fontSize: '36px',
          fontWeight: 'bold',
          opacity: skipIndicator.visible ? 1 : 0,
          transition: 'opacity 0.2s ease',
          pointerEvents: 'none',
          zIndex: 9999,
          display: 'flex',
          alignItems: 'center',
          gap: '15px',
          boxShadow: '0 8px 32px rgba(0,0,0,0.5)'
        }}
      >
        {skipIndicator.text.includes('-') ? (
           <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 17l-5-5 5-5M18 17l-5-5 5-5"/></svg>
        ) : (
           <svg xmlns="http://www.w3.org/2000/svg" width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M13 17l5-5-5-5M6 17l5-5-5-5"/></svg>
        )}
        {skipIndicator.text}
      </div>

      <video
        ref={recordingVideoRef}
        style={{
          display: 'none',
          position: 'absolute',
          top: 0,
          left: 0,
          width: '100%',
          height: '100%',
          objectFit: 'contain',
          zIndex: 2
        }}
        onTimeUpdate={handleRecordingTimeUpdate}
        onEnded={handleRecordingEnded}
        playsInline
        controls={false}
      />

      <div
        ref={containerRef}
        className="fs-timeline-container"
        style={{
          position: 'absolute',
          bottom: 0,
          left: 0,
          right: 0,
          zIndex: 20,
          padding: '6px 12px 10px 12px',
          background: 'linear-gradient(transparent, rgba(0,0,0,0.85))',
          opacity: visible ? 1 : 0,
          transition: 'opacity 0.3s ease',
          pointerEvents: visible ? 'auto' : 'none'
        }}
        onMouseEnter={() => {
          setVisible(true);
          if (hideTimerRef.current) clearTimeout(hideTimerRef.current);
        }}
        onMouseLeave={() => resetHideTimer()}
      >
        <div className="fs-timeline-shell">
          <div className="fs-timeline-side fs-timeline-side-left">
            <div className="fs-timeline-transport-group">
              <button type="button" className="fs-timeline-icon-button" title="Jump backward 60s" onClick={() => nudgeCursor(-60)}>
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M11 7.82v8.36L5.85 12 11 7.82Z"></path>
                  <path d="M18 7.82v8.36L12.85 12 18 7.82Z"></path>
                </svg>
              </button>
              <button type="button" className="fs-timeline-icon-button" title="Step backward 10s (Left Arrow)" onClick={() => nudgeCursor(-10)}>
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M15.5 7.82v8.36L10.35 12 15.5 7.82Z"></path>
                  <path d="M9.5 7.82v8.36L4.35 12 9.5 7.82Z"></path>
                </svg>
              </button>
              <button
                type="button"
                className="fs-timeline-icon-button fs-timeline-icon-button-primary"
                title={isPlayingRecording ? 'Pause or resume playback (Space)' : 'Live view active'}
                onClick={toggleRecordingPlayback}
                disabled={!isPlayingRecording}
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="currentColor">
                  {isPlayingRecording && !recordingPaused ? (
                    <>
                      <rect x="7" y="6" width="3.5" height="12" rx="1"></rect>
                      <rect x="13.5" y="6" width="3.5" height="12" rx="1"></rect>
                    </>
                  ) : (
                    <path d="M8 6v12l9-6-9-6Z"></path>
                  )}
                </svg>
              </button>
              <button type="button" className="fs-timeline-icon-button" title="Step forward 10s (Right Arrow)" onClick={() => nudgeCursor(10)}>
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M8.5 7.82v8.36L13.65 12 8.5 7.82Z"></path>
                  <path d="M14.5 7.82v8.36L19.65 12 14.5 7.82Z"></path>
                </svg>
              </button>
              <button type="button" className="fs-timeline-icon-button" title="Jump forward 60s" onClick={() => nudgeCursor(60)}>
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M6 7.82v8.36L11.15 12 6 7.82Z"></path>
                  <path d="M13 7.82v8.36L18.15 12 13 7.82Z"></path>
                </svg>
              </button>
            </div>
            <div className="fs-timeline-time-display">{formatPreciseTime(cursorTimestamp)}</div>
          </div>

          <div className="fs-timeline-main">
            <div className="fs-timeline-ruler">
              <div className="fs-timeline-day-label fs-timeline-day-label-left">
                {formatDayLabel(currentDate, 0)}
              </div>
              <div className="fs-timeline-day-label fs-timeline-day-label-right">
                {formatDayLabel(currentDate, visibleEndTimestamp >= dayEndTimestamp ? 1 : 0)}
              </div>
              {renderHourMarkers()}
            </div>
            <div className="fs-timeline-bar" onMouseDown={handleBarPointerDown}>
              <div className="fs-timeline-track"></div>
              {renderSegments()}
              {renderDetections()}
              {renderPositionIndicator()}
            </div>
          </div>

          <div className="fs-timeline-side fs-timeline-side-right">
            <div className="fs-timeline-status-group">
              <button type="button" className="fs-timeline-status-chip is-live">LIVE</button>
              <button type="button" className="fs-timeline-status-chip">SYNC</button>
            </div>
            <div className="fs-timeline-meta-actions fs-timeline-zoom-group">
              <button type="button" className="fs-timeline-icon-button" title="Zoom Out" onClick={zoomOut} disabled={zoomIndex === 0}>
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                  <circle cx="12" cy="12" r="8" strokeWidth="2"></circle>
                  <path d="M8 12h8" strokeWidth="2" strokeLinecap="round"></path>
                </svg>
              </button>
              <button type="button" className="fs-timeline-icon-button" title="Zoom In" onClick={zoomIn} disabled={zoomIndex === ZOOM_WINDOWS.length - 1}>
                <svg xmlns="http://www.w3.org/2000/svg" className="fs-timeline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                  <circle cx="12" cy="12" r="8" strokeWidth="2"></circle>
                  <path d="M8 12h8M12 8v8" strokeWidth="2" strokeLinecap="round"></path>
                </svg>
              </button>
              {isPlayingRecording ? (
                <button type="button" className="fs-timeline-status-action" onClick={returnToLive}>
                  Back to Live
                </button>
              ) : (
                <span className="fs-timeline-segment-count">{segments.length} segments</span>
              )}
            </div>
            <div className="fs-timeline-side-note">
              Window: {visibleWindowSeconds >= 3600
                ? `${Math.round(visibleWindowSeconds / 3600)}h`
                : visibleWindowSeconds >= 60
                  ? `${Math.round(visibleWindowSeconds / 60)}m`
                  : `${visibleWindowSeconds}s`}
            </div>
          </div>
        </div>
      </div>
    </>
  );
}
