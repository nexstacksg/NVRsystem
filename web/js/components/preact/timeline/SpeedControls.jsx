/**
 * Oneberry Speed Controls Component
 * Handles playback speed controls for the timeline
 */

import { useState, useEffect } from 'preact/hooks';
import { timelineState } from './TimelinePage.jsx';
import { showStatusMessage } from '../ToastContainer.jsx';

/**
 * SpeedControls component
 * @returns {JSX.Element} SpeedControls component
 */
export function SpeedControls() {
  // Local state
  const [currentSpeed, setCurrentSpeed] = useState(1.0);

  // Available speeds
  const speeds = [0.25, 0.5, 1.0, 1.5, 2.0, 4.0];

  // Subscribe to timeline state changes
  useEffect(() => {
    const unsubscribe = timelineState.subscribe(state => {
      setCurrentSpeed(state.playbackSpeed);
    });

    return () => unsubscribe();
  }, []);

  // Set playback speed
  const setPlaybackSpeed = (speed) => {
    // Update video playback rate
    const videoPlayer = document.querySelector('#video-player video');
    if (videoPlayer) {
      // Set the new playback rate
      videoPlayer.playbackRate = speed;
    }

    // Update timeline state
    timelineState.setState({ playbackSpeed: speed });

    // Show status message
    showStatusMessage(`Playback speed: ${speed}x`, 'info');
  };

  return (
    <div style={{
      display: 'flex',
      alignItems: 'center',
      gap: '6px'
    }}>
      <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--muted-foreground, #64748b)' }}>Speed:</span>
      <div style={{ display: 'flex', gap: '4px' }}>
        {speeds.map(speed => (
          <button
            key={`speed-${speed}`}
            className={`speed-btn ${speed === currentSpeed ? 'bg-primary text-primary-foreground' : 'bg-secondary text-secondary-foreground hover:bg-secondary/80'}`}
            style={{
              padding: '2px 8px',
              fontSize: '12px',
              borderRadius: '6px',
              border: 'none',
              cursor: 'pointer',
              fontWeight: 500,
              transition: 'all 0.15s ease'
            }}
            data-speed={speed}
            onClick={() => setPlaybackSpeed(speed)}
          >
            {speed === 1.0 ? '1×' : `${speed}×`}
          </button>
        ))}
      </div>
    </div>
  );
}
