/**
 * Detection overlay component for LiveView
 * Renders a canvas overlay for displaying detection boxes on video streams
 * Also renders a real-time motion grid overlay when detection model is 'motion'
 */
import { h } from 'preact';
import { useState, useEffect, useRef, useCallback } from 'preact/hooks';
import { showStatusMessage } from './ToastContainer.jsx';

import { forwardRef, useImperativeHandle } from 'preact/compat';

/**
 * DetectionOverlay component
 * @param {Object} props - Component props
 * @param {string} props.streamName - Name of the stream
 * @param {Object} props.videoRef - Reference to the video element
 * @param {boolean} props.enabled - Whether detection is enabled
 * @param {string} props.detectionModel - Detection model to use
 * @param {Object} ref - Forwarded ref
 * @returns {JSX.Element} DetectionOverlay component
 */
export const DetectionOverlay = forwardRef(({
  streamName,
  videoRef,
  enabled = false,
  detectionModel = null
}, ref) => {
  const [detections, setDetections] = useState([]);
  const [gridData, setGridData] = useState(null);
  const canvasRef = useRef(null);
  const intervalRef = useRef(null);
  const gridIntervalRef = useRef(null);
  const errorCountRef = useRef(0);
  const gridErrorCountRef = useRef(0);
  const currentIntervalRef = useRef(1000); // Start with 1 second polling interval
  const gridIntervalMs = 500; // 500ms for grid polling

  // Expose the canvas ref to parent components
  useImperativeHandle(ref, () => ({
    getCanvasRef: () => canvasRef,
    getDetections: () => detections
  }));

  // Function to draw bounding boxes and motion grid
  const drawDetectionBoxes = useCallback(() => {
    if (!canvasRef.current || !videoRef.current) return;

    const canvas = canvasRef.current;
    const videoElement = videoRef.current;
    const ctx = canvas.getContext('2d');

    // Set canvas dimensions to match the displayed video element
    canvas.width = videoElement.clientWidth;
    canvas.height = videoElement.clientHeight;

    // Clear previous drawings
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    // Get the actual video dimensions
    const videoWidth = videoElement.videoWidth;
    const videoHeight = videoElement.videoHeight;

    // If video dimensions aren't available yet, skip drawing
    if (!videoWidth || !videoHeight) {
      return;
    }

    // Calculate the scaling and positioning to maintain aspect ratio
    const videoAspect = videoWidth / videoHeight;
    const canvasAspect = canvas.width / canvas.height;

    let drawWidth, drawHeight, offsetX = 0, offsetY = 0;

    if (videoAspect > canvasAspect) {
      // Video is wider than canvas (letterboxing - black bars on top and bottom)
      drawWidth = canvas.width;
      drawHeight = canvas.width / videoAspect;
      offsetY = (canvas.height - drawHeight) / 2;
    } else {
      // Video is taller than canvas (pillarboxing - black bars on sides)
      drawHeight = canvas.height;
      drawWidth = canvas.height * videoAspect;
      offsetX = (canvas.width - drawWidth) / 2;
    }

    // Draw motion grid overlay
    // Always show grid lines when motion model is enabled, even without score data
    const isMotionModel = true; // Grid is drawn whenever gridData state exists
    const gridSize = (gridData && gridData.grid_size > 0) ? gridData.grid_size : 0;

    if (gridSize > 0) {
      const cellWidth = drawWidth / gridSize;
      const cellHeight = drawHeight / gridSize;

      // First: highlight cells with motion (green tint)
      if (gridData.scores && gridData.scores.length > 0) {
        for (let gy = 0; gy < gridSize; gy++) {
          for (let gx = 0; gx < gridSize; gx++) {
            const score = gridData.scores[gy * gridSize + gx];
            
            if (score > 0.005) {
              // Green highlight proportional to motion score
              const alpha = Math.min(0.5, Math.max(0.08, score * 6));
              
              const cellX = offsetX + gx * cellWidth;
              const cellY = offsetY + gy * cellHeight;

              ctx.fillStyle = `rgba(34, 197, 94, ${alpha})`;
              ctx.fillRect(cellX, cellY, cellWidth, cellHeight);

              // Draw cell border for active cells (brighter green)
              ctx.strokeStyle = `rgba(34, 197, 94, ${Math.min(0.8, alpha + 0.3)})`;
              ctx.lineWidth = 1;
              ctx.strokeRect(cellX, cellY, cellWidth, cellHeight);
            }
          }
        }
      }

      // Always draw grid lines across the entire frame
      ctx.strokeStyle = 'rgba(255, 255, 255, 0.2)';
      ctx.lineWidth = 0.5;

      for (let i = 1; i < gridSize; i++) {
        // Vertical lines
        const vx = offsetX + i * cellWidth;
        ctx.beginPath();
        ctx.moveTo(vx, offsetY);
        ctx.lineTo(vx, offsetY + drawHeight);
        ctx.stroke();

        // Horizontal lines
        const hy = offsetY + i * cellHeight;
        ctx.beginPath();
        ctx.moveTo(offsetX, hy);
        ctx.lineTo(offsetX + drawWidth, hy);
        ctx.stroke();
      }

      // Draw outer border of the grid
      ctx.strokeStyle = 'rgba(255, 255, 255, 0.3)';
      ctx.lineWidth = 1;
      ctx.strokeRect(offsetX, offsetY, drawWidth, drawHeight);
    }

    // Draw bounding box detections
    if (detections && detections.length > 0) {
      detections.forEach(detection => {
        // Calculate pixel coordinates based on normalized values (0-1)
        // and adjust for the actual display area
        const x = (detection.x * drawWidth) + offsetX;
        const y = (detection.y * drawHeight) + offsetY;
        const width = detection.width * drawWidth;
        const height = detection.height * drawHeight;

        // Draw bounding box
        ctx.strokeStyle = 'rgba(255, 0, 0, 0.8)';
        ctx.lineWidth = 3;
        ctx.strokeRect(x, y, width, height);

        // Draw label background
        const label = `${detection.label} (${Math.round(detection.confidence * 100)}%)`;
        ctx.font = '14px Arial';
        const textWidth = ctx.measureText(label).width;
        ctx.fillStyle = 'rgba(255, 0, 0, 0.7)';
        ctx.fillRect(x, y - 20, textWidth + 10, 20);

        // Draw label text
        ctx.fillStyle = 'white';
        ctx.fillText(label, x + 5, y - 5);
      });
    }
  }, [detections, gridData, videoRef]);

  // Poll for detections (bounding boxes from DB)
  const pollDetections = useCallback(() => {
    if (!videoRef.current || !videoRef.current.videoWidth) {
      return;
    }

    fetch(`/api/detection/results/${encodeURIComponent(streamName)}`)
      .then(response => {
        if (!response.ok) {
          throw new Error(`Failed to fetch detection results: ${response.status}`);
        }
        errorCountRef.current = 0;
        return response.json();
      })
      .then(data => {
        if (data && data.detections) {
          setDetections(data.detections);
        }
      })
      .catch(error => {
        console.error(`Error fetching detection results for ${streamName}:`, error);
        setDetections([]);

        errorCountRef.current++;
        if (errorCountRef.current > 3) {
          clearInterval(intervalRef.current);
          currentIntervalRef.current = Math.min(5000, currentIntervalRef.current * 2);
          console.log(`Reducing detection polling frequency to ${currentIntervalRef.current}ms due to errors`);
          intervalRef.current = setInterval(pollDetections, currentIntervalRef.current);
        }
      });
  }, [streamName, videoRef]);

  // Poll for motion grid scores (live from memory)
  const pollMotionGrid = useCallback(() => {
    if (!videoRef.current || !videoRef.current.videoWidth) {
      return;
    }

    fetch(`/api/detection/motion-grid/${encodeURIComponent(streamName)}`)
      .then(response => {
        if (!response.ok) {
          throw new Error(`Failed to fetch motion grid: ${response.status}`);
        }
        gridErrorCountRef.current = 0;
        return response.json();
      })
      .then(data => {
        if (data) {
          setGridData(data);
        }
      })
      .catch(error => {
        console.error(`Error fetching motion grid for ${streamName}:`, error);
        setGridData(null);

        gridErrorCountRef.current++;
        if (gridErrorCountRef.current > 5) {
          // Stop grid polling after too many errors
          if (gridIntervalRef.current) {
            clearInterval(gridIntervalRef.current);
            gridIntervalRef.current = null;
          }
          console.log(`Stopped motion grid polling for ${streamName} due to errors`);
        }
      });
  }, [streamName, videoRef]);

  // Start/stop detection polling based on enabled prop
  useEffect(() => {
    if (enabled && detectionModel && videoRef.current && canvasRef.current) {
      console.log(`Starting detection polling for stream ${streamName}`);

      if (intervalRef.current) {
        clearInterval(intervalRef.current);
      }

      intervalRef.current = setInterval(pollDetections, currentIntervalRef.current);

      return () => {
        console.log(`Cleaning up detection polling for stream ${streamName}`);
        if (intervalRef.current) {
          clearInterval(intervalRef.current);
          intervalRef.current = null;
        }
      };
    }

    if (intervalRef.current) {
      clearInterval(intervalRef.current);
      intervalRef.current = null;
    }
  }, [enabled, detectionModel, streamName, pollDetections, videoRef]);

  // Start/stop motion grid polling (only for 'motion' model)
  useEffect(() => {
    const isMotionModel = detectionModel && detectionModel.toLowerCase() === 'motion';
    
    if (enabled && isMotionModel && videoRef.current && canvasRef.current) {
      console.log(`Starting motion grid polling for stream ${streamName}`);

      if (gridIntervalRef.current) {
        clearInterval(gridIntervalRef.current);
      }

      gridErrorCountRef.current = 0;
      gridIntervalRef.current = setInterval(pollMotionGrid, gridIntervalMs);

      return () => {
        console.log(`Cleaning up motion grid polling for stream ${streamName}`);
        if (gridIntervalRef.current) {
          clearInterval(gridIntervalRef.current);
          gridIntervalRef.current = null;
        }
        setGridData(null);
      };
    }

    // Not a motion model or not enabled — clean up grid polling
    if (gridIntervalRef.current) {
      clearInterval(gridIntervalRef.current);
      gridIntervalRef.current = null;
    }
    setGridData(null);
  }, [enabled, detectionModel, streamName, pollMotionGrid, videoRef]);

  // Draw detections whenever they change
  useEffect(() => {
    drawDetectionBoxes();
  }, [detections, gridData, drawDetectionBoxes]);

  // Handle resize events to redraw detections
  useEffect(() => {
    const handleResize = () => {
      drawDetectionBoxes();
    };

    window.addEventListener('resize', handleResize);

    return () => {
      window.removeEventListener('resize', handleResize);
    };
  }, [drawDetectionBoxes]);

  // Handle fullscreen change events
  useEffect(() => {
    const handleFullscreenChange = () => {
      // Small delay to allow fullscreen to complete
      setTimeout(() => {
        drawDetectionBoxes();
      }, 100);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);

    return () => {
      document.removeEventListener('fullscreenchange', handleFullscreenChange);
    };
  }, [drawDetectionBoxes]);

  return (
    <canvas
      ref={canvasRef}
      className="detection-overlay"
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        width: '100%',
        height: '100%',
        pointerEvents: 'none',
        zIndex: 2
      }}
    />
  );
});

/**
 * Take a snapshot with detections
 * @param {Object} videoRef - Reference to the video element
 * @param {Object} canvasRef - Reference to the canvas element (from detectionOverlayRef.current.getCanvasRef())
 * @param {string} streamName - Name of the stream
 * @returns {Object} Canvas and filename for the snapshot
 */
export function takeSnapshotWithDetections(videoRef, canvasRef, streamName) {
  if (!videoRef.current || !canvasRef.current) {
    showStatusMessage('Cannot take snapshot: Video not available', 'error');
    return null;
  }

  const videoElement = videoRef.current;
  const canvasOverlay = canvasRef.current;

  // Create a combined canvas with video and detections
  const combinedCanvas = document.createElement('canvas');
  combinedCanvas.width = videoElement.videoWidth;
  combinedCanvas.height = videoElement.videoHeight;

  // Check if we have valid dimensions
  if (combinedCanvas.width === 0 || combinedCanvas.height === 0) {
    showStatusMessage('Cannot take snapshot: Video not loaded or has invalid dimensions', 'error');
    return null;
  }

  const ctx = combinedCanvas.getContext('2d');

  // Draw the video frame
  ctx.drawImage(videoElement, 0, 0, combinedCanvas.width, combinedCanvas.height);

  // Draw the detections from the overlay canvas
  if (canvasOverlay.width > 0 && canvasOverlay.height > 0) {
    ctx.drawImage(canvasOverlay, 0, 0, canvasOverlay.width, canvasOverlay.height,
                 0, 0, combinedCanvas.width, combinedCanvas.height);
  }

  // Generate a filename
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const fileName = `snapshot-${streamName.replace(/\s+/g, '-')}-${timestamp}.jpg`;

  return {
    canvas: combinedCanvas,
    fileName
  };
}
