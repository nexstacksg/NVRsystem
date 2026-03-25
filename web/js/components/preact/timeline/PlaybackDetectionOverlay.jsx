/**
 * PlaybackDetectionOverlay - Unifies backend bounding box detections 
 * and client-side pixel-diff motion grid for recording playback.
 */
import { h } from 'preact';
import { useState, useEffect, useRef, useCallback } from 'preact/hooks';
import { timelineState } from './TimelinePage.jsx';

const GRID_SIZE = 20;
const ANALYSIS_INTERVAL = 500; 
const MOTION_THRESHOLD = 25;  
const CELL_MOTION_THRESHOLD = 0.02; 
const ANALYSIS_SCALE = 0.25;  

export function PlaybackDetectionOverlay({ videoRef }) {
  const [streamConfig, setStreamConfig] = useState(null);
  const [detections, setDetections] = useState([]);
  const canvasRef = useRef(null);
  
  // Motion Grid state
  const analysisCanvasRef = useRef(null);
  const prevFrameRef = useRef(null);
  const gridScoresRef = useRef(new Float32Array(GRID_SIZE * GRID_SIZE));
  const gridIntervalRef = useRef(null);
  
  // Detections fetch state
  const lastFetchTimeRef = useRef(0);
  const detectionBufferRef = useRef([]);

  // Get stream config from timelineState
  useEffect(() => {
    const updateStream = () => {
      const streams = timelineState.streams || [];
      const currentStreamName = timelineState.selectedStream;
      const stream = streams.find(s => s.name === currentStreamName);
      setStreamConfig(stream || null);
    };

    updateStream();
    const unsubscribe = timelineState.subscribe(updateStream);
    return () => unsubscribe();
  }, []);

  // Poll for bounding boxes
  const pollDetections = useCallback(() => {
    if (!streamConfig || !streamConfig.detection_model || streamConfig.detection_model.toLowerCase() === 'none') {
      return;
    }

    const currentTime = timelineState.currentTime;
    if (!currentTime) return;

    // Fetch a 10-second window to avoid spamming the DB
    if (Math.abs(currentTime - lastFetchTimeRef.current) < 4) {
      return; // Still within cached window loosely
    }

    const start = Math.floor(currentTime - 2);
    const end = Math.floor(currentTime + 8);

    fetch(`/api/detection/results/${encodeURIComponent(streamConfig.name)}?start=${start}&end=${end}`)
      .then(res => res.ok ? res.json() : null)
      .then(data => {
        if (data && data.detections) {
          detectionBufferRef.current = data.detections;
          lastFetchTimeRef.current = currentTime;
        }
      })
      .catch(err => console.error("Error fetching playback detections:", err));
  }, [streamConfig]);

  useEffect(() => {
    if (!streamConfig || !streamConfig.detection_model || streamConfig.detection_model.toLowerCase() === 'none') {
      return;
    }
    const interval = setInterval(pollDetections, 1000);
    return () => clearInterval(interval);
  }, [streamConfig, pollDetections]);

  // Pixel-diff for motion grid
  const analyzeFrame = useCallback(() => {
    const video = videoRef.current;
    if (!video || video.paused || video.ended || !video.videoWidth) return;

    if (!analysisCanvasRef.current) {
      analysisCanvasRef.current = document.createElement('canvas');
    }

    const analysisCanvas = analysisCanvasRef.current;
    const aw = Math.floor(video.videoWidth * ANALYSIS_SCALE);
    const ah = Math.floor(video.videoHeight * ANALYSIS_SCALE);

    if (aw === 0 || ah === 0) return;

    analysisCanvas.width = aw;
    analysisCanvas.height = ah;

    const actx = analysisCanvas.getContext('2d', { willReadFrequently: true });
    actx.drawImage(video, 0, 0, aw, ah);
    const currentFrame = actx.getImageData(0, 0, aw, ah);
    const currentData = currentFrame.data;

    if (!prevFrameRef.current) {
      prevFrameRef.current = new Uint8ClampedArray(currentData);
      return;
    }

    const prevData = prevFrameRef.current;
    const scores = gridScoresRef.current;
    scores.fill(0);

    const cellW = aw / GRID_SIZE;
    const cellH = ah / GRID_SIZE;

    for (let gy = 0; gy < GRID_SIZE; gy++) {
      for (let gx = 0; gx < GRID_SIZE; gx++) {
        const startX = Math.floor(gx * cellW);
        const endX = Math.floor((gx + 1) * cellW);
        const startY = Math.floor(gy * cellH);
        const endY = Math.floor((gy + 1) * cellH);

        let changedPixels = 0;
        let totalPixels = 0;

        for (let y = startY; y < endY; y += 2) {
          for (let x = startX; x < endX; x += 2) {
            const idx = (y * aw + x) * 4;
            const currGray = (currentData[idx] * 77 + currentData[idx + 1] * 150 + currentData[idx + 2] * 29) >> 8;
            const prevGray = (prevData[idx] * 77 + prevData[idx + 1] * 150 + prevData[idx + 2] * 29) >> 8;
            const diff = Math.abs(currGray - prevGray);
            if (diff > MOTION_THRESHOLD) {
              changedPixels++;
            }
            totalPixels++;
          }
        }
        scores[gy * GRID_SIZE + gx] = totalPixels > 0 ? changedPixels / totalPixels : 0;
      }
    }
    prevFrameRef.current = new Uint8ClampedArray(currentData);
  }, [videoRef]);

  // Render loop
  const drawOverlay = useCallback(() => {
    const canvas = canvasRef.current;
    const video = videoRef.current;
    if (!canvas || !video) return;

    canvas.width = video.clientWidth;
    canvas.height = video.clientHeight;

    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const videoWidth = video.videoWidth;
    const videoHeight = video.videoHeight;
    if (!videoWidth || !videoHeight) return;

    const videoAspect = videoWidth / videoHeight;
    const canvasAspect = canvas.width / canvas.height;

    let drawWidth, drawHeight, offsetX = 0, offsetY = 0;

    if (videoAspect > canvasAspect) {
      drawWidth = canvas.width;
      drawHeight = canvas.width / videoAspect;
      offsetY = (canvas.height - drawHeight) / 2;
    } else {
      drawHeight = canvas.height;
      drawWidth = canvas.height * videoAspect;
      offsetX = (canvas.width - drawWidth) / 2;
    }

    // 1. Draw Motion Grid (only if stream explicitly supports it)
    const isMotionModel = streamConfig && streamConfig.detection_model && streamConfig.detection_model.toLowerCase() === 'motion';
    
    // We optionally can always draw grid lines, or only if motion model is enabled. LiveView only draws if motion model enabled.
    if (isMotionModel) {
      const cellWidth = drawWidth / GRID_SIZE;
      const cellHeight = drawHeight / GRID_SIZE;
      const scores = gridScoresRef.current;

      for (let gy = 0; gy < GRID_SIZE; gy++) {
        for (let gx = 0; gx < GRID_SIZE; gx++) {
          const score = scores[gy * GRID_SIZE + gx];
          if (score > CELL_MOTION_THRESHOLD) {
            const alpha = Math.min(0.5, Math.max(0.08, score * 6));
            const cellX = offsetX + gx * cellWidth;
            const cellY = offsetY + gy * cellHeight;

            ctx.fillStyle = `rgba(34, 197, 94, ${alpha})`;
            ctx.fillRect(cellX, cellY, cellWidth, cellHeight);

            ctx.strokeStyle = `rgba(34, 197, 94, ${Math.min(0.8, alpha + 0.3)})`;
            ctx.lineWidth = 1;
            ctx.strokeRect(cellX, cellY, cellWidth, cellHeight);
          }
        }
      }

      ctx.strokeStyle = 'rgba(255, 255, 255, 0.2)';
      ctx.lineWidth = 0.5;

      for (let i = 1; i < GRID_SIZE; i++) {
        const vx = offsetX + i * cellWidth;
        ctx.beginPath();
        ctx.moveTo(vx, offsetY);
        ctx.lineTo(vx, offsetY + drawHeight);
        ctx.stroke();

        const hy = offsetY + i * cellHeight;
        ctx.beginPath();
        ctx.moveTo(offsetX, hy);
        ctx.lineTo(offsetX + drawWidth, hy);
        ctx.stroke();
      }

      ctx.strokeStyle = 'rgba(255, 255, 255, 0.3)';
      ctx.lineWidth = 1;
      ctx.strokeRect(offsetX, offsetY, drawWidth, drawHeight);
    }

    // 2. Draw Bounding Boxes
    const currentTime = timelineState.currentTime;
    if (currentTime && detectionBufferRef.current.length > 0) {
      // Find detections within 1s of current time (could adjust window if desired limit)
      const visibleDetections = detectionBufferRef.current.filter(det => 
        Math.abs(det.timestamp - currentTime) < 1.0 // Within 1s
      );

      visibleDetections.forEach(detection => {
        const x = (detection.x * drawWidth) + offsetX;
        const y = (detection.y * drawHeight) + offsetY;
        const width = detection.width * drawWidth;
        const height = detection.height * drawHeight;

        ctx.strokeStyle = 'rgba(255, 0, 0, 0.8)';
        ctx.lineWidth = 3;
        ctx.strokeRect(x, y, width, height);

        const label = `${detection.label} (${Math.round(detection.confidence * 100)}%)`;
        ctx.font = '14px Arial';
        const textWidth = ctx.measureText(label).width;
        ctx.fillStyle = 'rgba(255, 0, 0, 0.7)';
        ctx.fillRect(x, y - 20, textWidth + 10, 20);

        ctx.fillStyle = 'white';
        ctx.fillText(label, x + 5, y - 5);
      });
    }
  }, [videoRef, streamConfig]);

  // Main effect loop
  useEffect(() => {
    let animFrame = null;
    
    // Pixel diffing analysis
    gridIntervalRef.current = setInterval(() => {
       // Only analyze if motion model is active to save CPU
       const isMotionModel = streamConfig && streamConfig.detection_model && streamConfig.detection_model.toLowerCase() === 'motion';
       if (isMotionModel) {
         analyzeFrame();
       }
    }, ANALYSIS_INTERVAL);

    // Render loop using requestAnimationFrame for smooth playback bounding boxes
    const renderLoop = () => {
      drawOverlay();
      animFrame = requestAnimationFrame(renderLoop);
    };
    animFrame = requestAnimationFrame(renderLoop);

    const handleResize = () => drawOverlay();
    window.addEventListener('resize', handleResize);
    document.addEventListener('fullscreenchange', handleResize);

    const video = videoRef.current;
    const handleSrcChange = () => {
      prevFrameRef.current = null;
      gridScoresRef.current.fill(0);
      drawOverlay();
    };

    if (video) {
      video.addEventListener('loadeddata', handleSrcChange);
      // We can also tie overlay to timeupdate, but requestAnimationFrame is smoother
    }

    return () => {
      clearInterval(gridIntervalRef.current);
      if (animFrame) cancelAnimationFrame(animFrame);
      window.removeEventListener('resize', handleResize);
      document.removeEventListener('fullscreenchange', handleResize);
      if (video) {
        video.removeEventListener('loadeddata', handleSrcChange);
      }
    };
  }, [analyzeFrame, drawOverlay, streamConfig, videoRef]);

  return (
    <canvas
      ref={canvasRef}
      className="playback-detection-overlay"
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
}
