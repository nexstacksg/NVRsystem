/**
 * FullscreenTimeline Component
 * A self-contained timeline bar for the camera fullscreen view.
 * Fetches recording segments independently and shows them as a compact
 * horizontal bar with hour markers, segment blocks, and a "now" indicator.
 * Clicking a segment switches from live to recorded playback.
 */

import { h } from 'preact';
import { useState, useEffect, useRef, useCallback } from 'preact/hooks';

/**
 * Format a Date to 'YYYY-MM-DD'
 */
function formatDateISO(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * Convert a timestamp (seconds) to fractional hour-of-day in local time.
 */
function timestampToHour(ts) {
  const d = new Date(ts * 1000);
  return d.getHours() + d.getMinutes() / 60 + d.getSeconds() / 3600;
}

/**
 * FullscreenTimeline component
 * @param {Object} props
 * @param {string} props.streamName - Name of the stream
 * @param {Object} props.videoRef - Ref to the live <video> element
 */
export function FullscreenTimeline({ streamName, videoRef }) {
  // ── state ──
  const [segments, setSegments] = useState([]);
  const [mergedSegments, setMergedSegments] = useState([]);
  const [nowHour, setNowHour] = useState(timestampToHour(Date.now() / 1000));
  const [visible, setVisible] = useState(true);
  const [isPlayingRecording, setIsPlayingRecording] = useState(false);
  const [playbackTime, setPlaybackTime] = useState(null); // timestamp of current playback position
  const [activeSegment, setActiveSegment] = useState(null); // segment being played

  // ── refs ──
  const containerRef = useRef(null);
  const hideTimerRef = useRef(null);
  const nowIntervalRef = useRef(null);
  const recordingVideoRef = useRef(null); // ref for our recording <video>

  // ── constants ──
  const START_HOUR = 0;
  const END_HOUR = 24;
  const HIDE_DELAY_MS = 4000;

  // ─────────────────────────────────────────────
  // Fetch segments for today
  // ─────────────────────────────────────────────
  useEffect(() => {
    if (!streamName) return;

    const today = new Date();
    const dateStr = formatDateISO(today);

    // Build start/end of day in local time, then convert to ISO for the API
    const [year, month, day] = dateStr.split('-').map(Number);
    const startDate = new Date(year, month - 1, day, 0, 0, 0, 0);
    const endDate = new Date(year, month - 1, day, 23, 59, 59, 999);

    const startTime = startDate.toISOString();
    const endTime = endDate.toISOString();

    const url = `/api/timeline/segments?stream=${encodeURIComponent(streamName)}&start=${encodeURIComponent(startTime)}&end=${encodeURIComponent(endTime)}`;

    console.log(`FullscreenTimeline: Fetching segments for ${streamName} on ${dateStr}`);

    fetch(url)
      .then(res => {
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return res.json();
      })
      .then(data => {
        const segs = data.segments || [];
        console.log(`FullscreenTimeline: Received ${segs.length} segments`);
        setSegments(segs);
      })
      .catch(err => {
        console.error('FullscreenTimeline: Error fetching segments:', err);
      });
  }, [streamName]);

  // ─────────────────────────────────────────────
  // Merge adjacent segments for cleaner rendering
  // ─────────────────────────────────────────────
  useEffect(() => {
    if (!segments || segments.length === 0) {
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
        // merge
        current.end_timestamp = Math.max(current.end_timestamp, seg.end_timestamp);
        current.originalSegments.push(seg);
      } else {
        merged.push(current);
        current = { ...seg, originalSegments: [seg] };
      }
    }
    merged.push(current);
    setMergedSegments(merged);
  }, [segments]);

  // ─────────────────────────────────────────────
  // Update "now" marker every 30 seconds
  // ─────────────────────────────────────────────
  useEffect(() => {
    const update = () => setNowHour(timestampToHour(Date.now() / 1000));
    nowIntervalRef.current = setInterval(update, 30000);
    return () => clearInterval(nowIntervalRef.current);
  }, []);

  // ─────────────────────────────────────────────
  // Auto-hide logic
  // ─────────────────────────────────────────────
  const resetHideTimer = useCallback(() => {
    setVisible(true);
    if (hideTimerRef.current) clearTimeout(hideTimerRef.current);
    hideTimerRef.current = setTimeout(() => setVisible(false), HIDE_DELAY_MS);
  }, []);

  useEffect(() => {
    // Listen for mouse movement on the fullscreen element
    const handleMouseMove = () => resetHideTimer();

    document.addEventListener('mousemove', handleMouseMove);
    resetHideTimer(); // start initial timer

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      if (hideTimerRef.current) clearTimeout(hideTimerRef.current);
    };
  }, [resetHideTimer]);

  // ─────────────────────────────────────────────
  // Handle clicking on a segment to play recording
  // ─────────────────────────────────────────────
  const handleTimelineClick = useCallback((e) => {
    const bar = containerRef.current?.querySelector('.fs-timeline-bar');
    if (!bar) return;

    const rect = bar.getBoundingClientRect();
    const clickX = e.clientX - rect.left;
    const pct = clickX / rect.width;
    const clickHour = START_HOUR + pct * (END_HOUR - START_HOUR);

    // Convert click hour to timestamp
    const today = new Date();
    const [year, month, day] = [today.getFullYear(), today.getMonth(), today.getDate()];
    const clickDate = new Date(year, month, day, 0, 0, 0, 0);
    const clickTimestamp = clickDate.getTime() / 1000 + clickHour * 3600;

    // Find which original segment contains this timestamp
    let foundSeg = null;
    for (const seg of segments) {
      if (clickTimestamp >= seg.start_timestamp && clickTimestamp <= seg.end_timestamp) {
        foundSeg = seg;
        break;
      }
    }

    if (!foundSeg) {
      // Find closest segment
      let closest = null;
      let minDist = Infinity;
      for (const seg of segments) {
        const mid = (seg.start_timestamp + seg.end_timestamp) / 2;
        const dist = Math.abs(clickTimestamp - mid);
        if (dist < minDist) {
          minDist = dist;
          closest = seg;
        }
      }
      if (closest && minDist < 1800) { // within 30 minutes
        foundSeg = closest;
      }
    }

    if (!foundSeg) {
      console.log('FullscreenTimeline: No segment found near clicked position');
      return;
    }

    console.log(`FullscreenTimeline: Playing segment ${foundSeg.id} from timestamp ${clickTimestamp}`);

    // Calculate seek time within the segment
    const seekTime = Math.max(0, clickTimestamp - foundSeg.start_timestamp);

    // Hide the live video and show recorded playback
    playRecording(foundSeg, seekTime);
  }, [segments]);

  // ─────────────────────────────────────────────
  // Play a recording segment
  // ─────────────────────────────────────────────
  const playRecording = useCallback((segment, seekTime = 0) => {
    setIsPlayingRecording(true);
    setActiveSegment(segment);
    setPlaybackTime(segment.start_timestamp + seekTime);

    // Pause the live video
    if (videoRef?.current) {
      videoRef.current.style.display = 'none';
    }

    // Load recording into our recording video element
    const recVideo = recordingVideoRef.current;
    if (!recVideo) return;

    recVideo.style.display = 'block';

    const url = `/api/recordings/play/${segment.id}?t=${Date.now()}`;
    recVideo.src = url;

    const onMetadata = () => {
      recVideo.currentTime = Math.min(seekTime, recVideo.duration || seekTime);
      recVideo.play().catch(err => console.error('FullscreenTimeline: playback error:', err));
      recVideo.removeEventListener('loadedmetadata', onMetadata);
    };

    recVideo.addEventListener('loadedmetadata', onMetadata);
    recVideo.load();
  }, [videoRef]);

  // ─────────────────────────────────────────────
  // Return to live view
  // ─────────────────────────────────────────────
  const returnToLive = useCallback(() => {
    setIsPlayingRecording(false);
    setActiveSegment(null);
    setPlaybackTime(null);

    // Stop recording video
    const recVideo = recordingVideoRef.current;
    if (recVideo) {
      recVideo.pause();
      recVideo.removeAttribute('src');
      recVideo.load();
      recVideo.style.display = 'none';
    }

    // Show live video again
    if (videoRef?.current) {
      videoRef.current.style.display = 'block';
    }
  }, [videoRef]);

  // ─────────────────────────────────────────────
  // Track playback time while recording plays
  // ─────────────────────────────────────────────
  const handleRecordingTimeUpdate = useCallback(() => {
    const recVideo = recordingVideoRef.current;
    if (!recVideo || !activeSegment) return;
    setPlaybackTime(activeSegment.start_timestamp + recVideo.currentTime);
  }, [activeSegment]);

  // ─────────────────────────────────────────────
  // Handle recording ended — try next segment or go back to live
  // ─────────────────────────────────────────────
  const handleRecordingEnded = useCallback(() => {
    if (!activeSegment) {
      returnToLive();
      return;
    }

    // Find the next segment
    const currentIdx = segments.findIndex(s => s.id === activeSegment.id);
    if (currentIdx >= 0 && currentIdx < segments.length - 1) {
      const nextSeg = segments[currentIdx + 1];
      playRecording(nextSeg, 0);
    } else {
      returnToLive();
    }
  }, [activeSegment, segments, playRecording, returnToLive]);

  // ─────────────────────────────────────────────
  // Render helper: hour markers
  // ─────────────────────────────────────────────
  const renderHourMarkers = () => {
    const markers = [];
    for (let h = 0; h <= 24; h++) {
      const pct = ((h - START_HOUR) / (END_HOUR - START_HOUR)) * 100;
      // Major tick
      markers.push(
        <div
          key={`tick-${h}`}
          style={{
            position: 'absolute',
            left: `${pct}%`,
            top: 0,
            width: '1px',
            height: '10px',
            backgroundColor: 'rgba(255,255,255,0.4)'
          }}
        />
      );
      // Label (skip if too close to the edge)
      if (h < 24) {
        markers.push(
          <div
            key={`lbl-${h}`}
            style={{
              position: 'absolute',
              left: `${pct}%`,
              top: '0px',
              fontSize: '9px',
              color: 'rgba(255,255,255,0.6)',
              transform: 'translateX(2px)',
              userSelect: 'none',
              whiteSpace: 'nowrap'
            }}
          >
            {String(h).padStart(2, '0')}:00
          </div>
        );
      }
    }
    return markers;
  };

  // ─────────────────────────────────────────────
  // Render helper: segment blocks
  // ─────────────────────────────────────────────
  const renderSegments = () => {
    return mergedSegments.map((seg, i) => {
      const startH = timestampToHour(seg.start_timestamp);
      const endH = timestampToHour(seg.end_timestamp);

      const leftPct = ((startH - START_HOUR) / (END_HOUR - START_HOUR)) * 100;
      const widthPct = ((endH - startH) / (END_HOUR - START_HOUR)) * 100;

      return (
        <div
          key={`seg-${i}`}
          style={{
            position: 'absolute',
            left: `${leftPct}%`,
            width: `${Math.max(widthPct, 0.3)}%`,
            top: '14px',
            height: '16px',
            backgroundColor: seg.has_detection ? 'rgba(239, 68, 68, 0.8)' : 'rgba(156, 163, 175, 0.8)',
            borderRadius: '2px',
            cursor: 'pointer',
            transition: 'background-color 0.15s ease'
          }}
          title={`${new Date(seg.start_timestamp * 1000).toLocaleTimeString()} - ${new Date(seg.end_timestamp * 1000).toLocaleTimeString()}`}
          onMouseOver={(e) => {
            e.currentTarget.style.backgroundColor = seg.has_detection
              ? 'rgba(239, 68, 68, 1)'
              : 'rgba(156, 163, 175, 1)';
          }}
          onMouseOut={(e) => {
            e.currentTarget.style.backgroundColor = seg.has_detection
              ? 'rgba(239, 68, 68, 0.8)'
              : 'rgba(156, 163, 175, 0.8)';
          }}
        />
      );
    });
  };

  // ─────────────────────────────────────────────
  // Render helper: "now" / playback position indicator
  // ─────────────────────────────────────────────
  const renderPositionIndicator = () => {
    const indicatorHour = isPlayingRecording && playbackTime
      ? timestampToHour(playbackTime)
      : nowHour;

    const pct = ((indicatorHour - START_HOUR) / (END_HOUR - START_HOUR)) * 100;

    if (pct < 0 || pct > 100) return null;

    return (
      <div
        style={{
          position: 'absolute',
          left: `${pct}%`,
          top: '12px',
          width: '2px',
          height: '20px',
          backgroundColor: '#f97316',
          zIndex: 10,
          pointerEvents: 'none',
          boxShadow: '0 0 4px rgba(249, 115, 22, 0.6)'
        }}
      />
    );
  };

  // ─────────────────────────────────────────────
  // Format time for display
  // ─────────────────────────────────────────────
  const formatTime = (hour) => {
    const h = Math.floor(hour);
    const m = Math.floor((hour - h) * 60);
    const s = Math.floor(((hour - h) * 60 - m) * 60);
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  };

  const displayHour = isPlayingRecording && playbackTime
    ? timestampToHour(playbackTime)
    : nowHour;

  // ─────────────────────────────────────────────
  // Main render
  // ─────────────────────────────────────────────
  return (
    <>
      {/* Recording video element (hidden by default) */}
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

      {/* Timeline bar container */}
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
        {/* Top row: status + time display + back-to-live button */}
        <div style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: '4px'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {/* Live / Recording indicator */}
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '4px',
              fontSize: '11px',
              color: isPlayingRecording ? '#f97316' : '#22c55e',
              fontWeight: 'bold',
              textTransform: 'uppercase',
              letterSpacing: '0.5px'
            }}>
              <div style={{
                width: '6px',
                height: '6px',
                borderRadius: '50%',
                backgroundColor: isPlayingRecording ? '#f97316' : '#22c55e',
                boxShadow: isPlayingRecording
                  ? '0 0 6px rgba(249, 115, 22, 0.6)'
                  : '0 0 6px rgba(34, 197, 94, 0.6)'
              }} />
              {isPlayingRecording ? 'Recording' : 'Live'}
            </div>

            {/* Time display */}
            <span style={{
              fontFamily: 'monospace',
              fontSize: '12px',
              color: 'rgba(255,255,255,0.8)',
              backgroundColor: 'rgba(0,0,0,0.4)',
              padding: '1px 6px',
              borderRadius: '3px'
            }}>
              {formatTime(displayHour)}
            </span>
          </div>

          {/* Back to Live button */}
          {isPlayingRecording && (
            <button
              onClick={returnToLive}
              style={{
                fontSize: '11px',
                color: 'white',
                backgroundColor: 'rgba(34, 197, 94, 0.8)',
                border: 'none',
                padding: '2px 10px',
                borderRadius: '3px',
                cursor: 'pointer',
                fontWeight: 'bold',
                transition: 'background-color 0.15s ease'
              }}
              onMouseOver={(e) => e.currentTarget.style.backgroundColor = 'rgba(34, 197, 94, 1)'}
              onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'rgba(34, 197, 94, 0.8)'}
            >
              ● Back to Live
            </button>
          )}

          {/* Segment count */}
          {segments.length > 0 && (
            <span style={{
              fontSize: '10px',
              color: 'rgba(255,255,255,0.4)'
            }}>
              {segments.length} segments
            </span>
          )}
        </div>

        {/* Timeline bar */}
        <div
          className="fs-timeline-bar"
          onClick={handleTimelineClick}
          style={{
            position: 'relative',
            width: '100%',
            height: '34px',
            backgroundColor: 'rgba(255,255,255,0.08)',
            borderRadius: '4px',
            cursor: 'pointer',
            overflow: 'hidden'
          }}
        >
          {renderHourMarkers()}
          {renderSegments()}
          {renderPositionIndicator()}
        </div>
      </div>
    </>
  );
}
