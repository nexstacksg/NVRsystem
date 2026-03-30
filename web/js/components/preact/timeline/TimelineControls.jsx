/**
 * Oneberry Timeline Controls Component
 * Handles play/pause and zoom controls for the timeline
 */

import { useState, useEffect } from 'preact/hooks';
import { timelineState } from './TimelinePage.jsx';
import { showStatusMessage } from '../ToastContainer.jsx';

/**
 * TimelineControls component
 * @returns {JSX.Element} TimelineControls component
 */
export function TimelineControls({ compact = false }) {
  // Local state
  const [isPlaying, setIsPlaying] = useState(false);
  const [zoomLevel, setZoomLevel] = useState(1);

  // Subscribe to timeline state changes
  useEffect(() => {
    const unsubscribe = timelineState.subscribe(state => {
      setIsPlaying(state.isPlaying);
      setZoomLevel(state.zoomLevel);
    });

    return () => unsubscribe();
  }, []);

  // Toggle playback (play/pause)
  const togglePlayback = () => {
    console.log('TimelineControls: togglePlayback called');
    console.log('TimelineControls: Current state before toggle:', {
      isPlaying,
      currentTime: timelineState.currentTime,
      currentSegmentIndex: timelineState.currentSegmentIndex,
      segmentsCount: timelineState.timelineSegments ? timelineState.timelineSegments.length : 0
    });

    if (isPlaying) {
      pausePlayback();
    } else {
      resumePlayback();
    }
  };

  // Pause playback
  const pausePlayback = () => {
    timelineState.setState({ isPlaying: false });

    // If there's a video player, pause it
    const videoPlayer = document.querySelector('#video-player video');
    if (videoPlayer) {
      videoPlayer.pause();
    }
  };

  // Resume playback
  const resumePlayback = () => {
    // If no segments, show message and return
    if (!timelineState.timelineSegments || timelineState.timelineSegments.length === 0) {
      showStatusMessage('No recordings to play', 'warning');
      return;
    }

    console.log('TimelineControls: resumePlayback called');
    console.log('TimelineControls: Current state:', {
      segments: timelineState.timelineSegments.length,
      currentSegmentIndex: timelineState.currentSegmentIndex,
      currentTime: timelineState.currentTime,
      selectedDate: timelineState.selectedDate
    });

    // SIMPLIFIED APPROACH: Directly use the current time to find the appropriate segment
    let segmentToPlay = null;
    let segmentIndex = -1;
    let relativeTime = 0;

    // If we have a current time, find the segment that contains it
    if (timelineState.currentTime !== null) {
      console.log('TimelineControls: Using current time to find segment:', timelineState.currentTime);

      // First try to find a segment that contains the current time
      for (let i = 0; i < timelineState.timelineSegments.length; i++) {
        const segment = timelineState.timelineSegments[i];
        if (timelineState.currentTime >= segment.start_timestamp &&
            timelineState.currentTime <= segment.end_timestamp) {
          segmentToPlay = segment;
          segmentIndex = i;
          relativeTime = timelineState.currentTime - segment.start_timestamp;
          console.log(`TimelineControls: Found segment ${i} containing current time, relative time: ${relativeTime}s`);
          break;
        }
      }

      // If no exact match, find the closest segment
      if (!segmentToPlay) {
        let closestIndex = 0;
        let minDistance = Infinity;

        for (let i = 0; i < timelineState.timelineSegments.length; i++) {
          const segment = timelineState.timelineSegments[i];
          const midpoint = (segment.start_timestamp + segment.end_timestamp) / 2;
          const distance = Math.abs(timelineState.currentTime - midpoint);

          if (distance < minDistance) {
            minDistance = distance;
            closestIndex = i;
          }
        }

        segmentToPlay = timelineState.timelineSegments[closestIndex];
        segmentIndex = closestIndex;

        // Calculate the appropriate relative time based on the cursor position
        if (timelineState.currentTime < segmentToPlay.start_timestamp) {
          // If cursor is before the segment, start at the beginning
          relativeTime = 0;
        } else if (timelineState.currentTime > segmentToPlay.end_timestamp) {
          // If cursor is after the segment, start at the end
          relativeTime = segmentToPlay.end_timestamp - segmentToPlay.start_timestamp;
        } else {
          // If cursor is within the segment's time range but not exactly matching,
          // calculate the relative position within the segment
          relativeTime = timelineState.currentTime - segmentToPlay.start_timestamp;
        }

        console.log(`TimelineControls: Using closest segment ${closestIndex}, relative time: ${relativeTime}s`);
      }
    }
    // If no current time but we have a valid segment index, use that
    else if (timelineState.currentSegmentIndex >= 0 &&
             timelineState.currentSegmentIndex < timelineState.timelineSegments.length) {
      segmentIndex = timelineState.currentSegmentIndex;
      segmentToPlay = timelineState.timelineSegments[segmentIndex];
      relativeTime = 0;
      console.log(`TimelineControls: Using current segment index ${segmentIndex}`);
    }
    // Fall back to the first segment
    else {
      segmentIndex = 0;
      segmentToPlay = timelineState.timelineSegments[0];
      relativeTime = 0;
      console.log('TimelineControls: Falling back to first segment');
    }

    // DIRECT APPROACH: Manually load and play the video
    console.log(`TimelineControls: Playing segment ${segmentIndex} (ID: ${segmentToPlay.id}) at time ${relativeTime}s`);

    // First update the state
    timelineState.currentSegmentIndex = segmentIndex;

    // Check if we should preserve the cursor position
    if (timelineState.preserveCursorPosition) {
      // If the cursor position is locked, don't change the current time at all
      console.log(`TimelineControls: Cursor position is locked, keeping current time ${timelineState.currentTime}`);
    } else {
      // Make sure we preserve the current time if it's within the segment
      // This ensures the cursor stays where the user positioned it
      if (timelineState.currentTime !== null &&
          timelineState.currentTime >= segmentToPlay.start_timestamp &&
          timelineState.currentTime <= segmentToPlay.end_timestamp) {
        // Keep the current time as is - don't reset it
        console.log(`TimelineControls: Keeping current time ${timelineState.currentTime} within segment`);
      } else {
        // Otherwise, set the time based on the calculated relative time
        timelineState.currentTime = segmentToPlay.start_timestamp + relativeTime;
        console.log(`TimelineControls: Setting current time to ${timelineState.currentTime} (segment start + ${relativeTime}s)`);
      }
    }

    timelineState.isPlaying = true;
    timelineState.directVideoControl = true; // Set flag to prevent TimelinePlayer interference

    // Notify listeners
    timelineState.setState({});

    // Keep the directVideoControl flag active longer to ensure no interference
    const resetDirectControl = () => {
      console.log('TimelineControls: Resetting directVideoControl flag');
      timelineState.directVideoControl = false;
      timelineState.setState({});
    };

    // Reset the flag after a longer delay
    setTimeout(resetDirectControl, 3000);

    // Now directly control the video element
    const videoElement = document.querySelector('#video-player video');
    if (videoElement) {
      // Pause any current playback
      videoElement.pause();

      // Set up event listener for when metadata is loaded
      const handleMetadataLoaded = () => {
        console.log(`TimelineControls: Video metadata loaded, setting time to ${relativeTime}s`);

        try {
          // Log the video duration
          console.log('TimelineControls: Video metadata', {
            duration: videoElement.duration,
            width: videoElement.videoWidth,
            height: videoElement.videoHeight,
            segment: segmentToPlay.id,
            segmentDuration: segmentToPlay.end_timestamp - segmentToPlay.start_timestamp
          });

          // Set the current time based on the cursor position
          let timeToSet = relativeTime;

          // Check if we should preserve the cursor position
          if (timelineState.preserveCursorPosition && timelineState.currentTime !== null) {
            // Calculate the relative time within the segment
            if (timelineState.currentTime >= segmentToPlay.start_timestamp &&
                timelineState.currentTime <= segmentToPlay.end_timestamp) {
              // If the cursor is within this segment, use its position
              timeToSet = timelineState.currentTime - segmentToPlay.start_timestamp;
              console.log(`TimelineControls: Using locked cursor position for playback: ${timeToSet}s`);
            } else {
              // If the cursor is outside this segment but we want to preserve its position,
              // use the beginning or end of the segment based on which is closer
              const distanceToStart = Math.abs(timelineState.currentTime - segmentToPlay.start_timestamp);
              const distanceToEnd = Math.abs(timelineState.currentTime - segmentToPlay.end_timestamp);

              if (distanceToStart <= distanceToEnd) {
                timeToSet = 0; // Use start of segment
                console.log(`TimelineControls: Cursor outside segment, using start of segment`);
              } else {
                timeToSet = segmentToPlay.end_timestamp - segmentToPlay.start_timestamp; // Use end of segment
                console.log(`TimelineControls: Cursor outside segment, using end of segment`);
              }
            }
          } else if (timelineState.currentTime !== null &&
                     timelineState.currentTime >= segmentToPlay.start_timestamp &&
                     timelineState.currentTime <= segmentToPlay.end_timestamp) {
            // If not preserving cursor position but the current time is within this segment,
            // still use the relative position from the cursor
            timeToSet = timelineState.currentTime - segmentToPlay.start_timestamp;
            console.log(`TimelineControls: Using cursor position for playback: ${timeToSet}s`);
          }

          // Ensure the time is within valid bounds
          let validTime = Math.max(0, Math.min(timeToSet, videoElement.duration || 0));

          // If we're very close to the start of the segment but not exactly at the start,
          // add a small offset to prevent snapping to the start
          if (validTime > 0 && validTime < 1.0) {
            // Ensure we're at least 1 second into the segment
            validTime = 1.0;
            console.log(`TimelineControls: Adjusting video time to ${validTime}s to prevent snapping to segment start`);
          }

          console.log(`TimelineControls: Setting video time to ${validTime}s`);
          videoElement.currentTime = validTime;

          // Start playback with a small delay to avoid conflicts
          setTimeout(() => {
            if (timelineState.isPlaying) {
              console.log('TimelineControls: Starting video playback');
              videoElement.play().then(() => {
                console.log('TimelineControls: Video playback started successfully');

                // Set up multiple checks to ensure playback continues
                const checkPlayback = (attempt = 1) => {
                  if (attempt > 5) return; // Limit to 5 attempts

                  setTimeout(() => {
                    if (videoElement.paused && timelineState.isPlaying) {
                      console.log(`TimelineControls: Video paused unexpectedly (attempt ${attempt}), trying to resume`);
                      videoElement.play().catch(e => {
                        console.error(`Error resuming video (attempt ${attempt}):`, e);
                      });

                      // Try again after a delay
                      checkPlayback(attempt + 1);
                    }
                  }, 500 * attempt); // Increasing delays between attempts
                };

                // Start the playback checks
                checkPlayback();

                // Set up a check to ensure the video plays for the full segment duration
                const segmentDuration = segmentToPlay.end_timestamp - segmentToPlay.start_timestamp;
                console.log(`TimelineControls: Segment duration: ${segmentDuration}s, video duration: ${videoElement.duration}s`);

                // If the video duration is significantly shorter than the segment duration,
                // we need to ensure the video plays for the full segment duration
                if (videoElement.duration < segmentDuration - 1) { // Allow 1 second tolerance
                  console.log('TimelineControls: Video duration is shorter than segment duration, will monitor playback');

                  // Monitor playback to ensure it continues for the full segment duration
                  const monitorInterval = setInterval(() => {
                    if (!timelineState.isPlaying || !videoElement) {
                      clearInterval(monitorInterval);
                      return;
                    }

                    // If we're near the end of the video but not the end of the segment,
                    // reset to the beginning and continue playing
                    if (videoElement.currentTime > videoElement.duration - 0.5 &&
                        relativeTime + videoElement.currentTime < segmentDuration) {
                      console.log('TimelineControls: Reached end of video but not end of segment, restarting video');
                      videoElement.currentTime = 0;
                      videoElement.play().catch(e => {
                        console.error('Error restarting video:', e);
                      });
                    }
                  }, 500);
                }
              }).catch(e => {
                console.error('Error playing video:', e);
                showStatusMessage('Error playing video: ' + e.message, 'error');
              });
            }
          }, 100);
        } catch (error) {
          console.error('TimelineControls: Error in handleMetadataLoaded:', error);
        } finally {
          // Remove the event listener
          videoElement.removeEventListener('loadedmetadata', handleMetadataLoaded);
        }
      };

      // Add the event listener
      videoElement.addEventListener('loadedmetadata', handleMetadataLoaded);

      // Set the new source
      console.log(`TimelineControls: Loading video from segment ${segmentToPlay.id}`);
      videoElement.src = `/api/recordings/play/${segmentToPlay.id}?t=${Date.now()}`;
      videoElement.load();
    } else {
      console.error('TimelineControls: No video element found');
      showStatusMessage('Error: Video player not found', 'error');
    }
  };

  // Play a specific segment
  const playSegment = (index) => {
    // This function will be implemented in the TimelinePlayer component
    // Here we just update the state
    timelineState.setState({
      currentSegmentIndex: index,
      isPlaying: true
    });
  };

  // Zoom in on timeline
  const zoomIn = () => {
    if (zoomLevel < 8) {
      const newZoomLevel = zoomLevel * 2;
      timelineState.setState({ zoomLevel: newZoomLevel });
      showStatusMessage(`Zoomed in: ${24 / newZoomLevel} hours view`, 'info');
    }
  };

  // Zoom out on timeline
  const zoomOut = () => {
    if (zoomLevel > 1) {
      const newZoomLevel = zoomLevel / 2;
      timelineState.setState({ zoomLevel: newZoomLevel });
      showStatusMessage(`Zoomed out: ${24 / newZoomLevel} hours view`, 'info');
    }
  };

  if (compact) {
    return (
      <div className="timeline-controls timeline-controls-compact">
        <div className="timeline-transport-group">
          <button type="button" className="timeline-icon-button" title="Jump backward 10 seconds">
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="currentColor">
              <path d="M11 7.82v8.36L5.85 12 11 7.82Z"></path>
              <path d="M18 7.82v8.36L12.85 12 18 7.82Z"></path>
            </svg>
          </button>
          <button type="button" className="timeline-icon-button" title="Back">
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="currentColor">
              <path d="M15.5 7.82v8.36L10.35 12 15.5 7.82Z"></path>
              <path d="M9.5 7.82v8.36L4.35 12 9.5 7.82Z"></path>
            </svg>
          </button>
          <button
            id="play-button"
            type="button"
            className="timeline-icon-button timeline-icon-button-primary"
            onClick={togglePlayback}
            title={isPlaying ? 'Pause' : 'Play from current position'}
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="currentColor">
              {isPlaying ? (
                <>
                  <rect x="7" y="6" width="3.5" height="12" rx="1"></rect>
                  <rect x="13.5" y="6" width="3.5" height="12" rx="1"></rect>
                </>
              ) : (
                <path d="M8 6v12l9-6-9-6Z"></path>
              )}
            </svg>
          </button>
          <button type="button" className="timeline-icon-button" title="Forward">
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="currentColor">
              <path d="M8.5 7.82v8.36L13.65 12 8.5 7.82Z"></path>
              <path d="M14.5 7.82v8.36L19.65 12 14.5 7.82Z"></path>
            </svg>
          </button>
          <button type="button" className="timeline-icon-button" title="Jump forward 10 seconds">
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="currentColor">
              <path d="M6 7.82v8.36L11.15 12 6 7.82Z"></path>
              <path d="M13 7.82v8.36L18.15 12 13 7.82Z"></path>
            </svg>
          </button>
        </div>
        <div className="timeline-zoom-group">
          <button
            id="zoom-out-button"
            type="button"
            className="timeline-icon-button"
            onClick={zoomOut}
            title="Zoom Out"
            disabled={zoomLevel <= 1}
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
              <circle cx="12" cy="12" r="8" strokeWidth="2"></circle>
              <path d="M8 12h8" strokeWidth="2" strokeLinecap="round"></path>
            </svg>
          </button>
          <button
            id="zoom-in-button"
            type="button"
            className="timeline-icon-button"
            onClick={zoomIn}
            title="Zoom In"
            disabled={zoomLevel >= 8}
          >
            <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
              <circle cx="12" cy="12" r="8" strokeWidth="2"></circle>
              <path d="M8 12h8M12 8v8" strokeWidth="2" strokeLinecap="round"></path>
            </svg>
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="timeline-controls">
      <div className="timeline-controls-main">
        <button
          id="play-button"
          className="timeline-icon-button timeline-icon-button-primary"
          onClick={togglePlayback}
          title={isPlaying ? 'Pause' : 'Play from current position'}
        >
          <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="currentColor">
            {isPlaying ? (
              <>
                <rect x="7" y="6" width="3.5" height="12" rx="1"></rect>
                <rect x="13.5" y="6" width="3.5" height="12" rx="1"></rect>
              </>
            ) : (
              <path d="M8 6v12l9-6-9-6Z"></path>
            )}
          </svg>
        </button>
        <span className="timeline-controls-label">Play from current position</span>
      </div>

      <div className="timeline-zoom-group">
        <button
          id="zoom-out-button"
          className="timeline-icon-button"
          onClick={zoomOut}
          title="Zoom Out (Show more time)"
          disabled={zoomLevel <= 1}
        >
          <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
            <circle cx="12" cy="12" r="8" strokeWidth="2"></circle>
            <path d="M8 12h8" strokeWidth="2" strokeLinecap="round"></path>
          </svg>
        </button>
        <button
          id="zoom-in-button"
          className="timeline-icon-button"
          onClick={zoomIn}
          title="Zoom In (Show less time)"
          disabled={zoomLevel >= 8}
        >
          <svg xmlns="http://www.w3.org/2000/svg" className="timeline-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
            <circle cx="12" cy="12" r="8" strokeWidth="2"></circle>
            <path d="M8 12h8M12 8v8" strokeWidth="2" strokeLinecap="round"></path>
          </svg>
        </button>
      </div>
    </div>
  );
}
