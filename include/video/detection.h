#ifndef DETECTION_H
#define DETECTION_H

#include <stdbool.h>
#include "video/detection_result.h"
#include "video/detection_model.h"
#include "video/sod_detection.h"
#include "video/onvif_detection.h"

/**
 * Initialize the detection system
 * 
 * @return 0 on success, non-zero on failure
 */
int init_detection_system(void);

/**
 * Shutdown the detection system
 */
void shutdown_detection_system(void);

/**
 * Run detection on a frame
 * 
 * @param model Detection model handle
 * @param frame_data Pointer to the frame data (RGB)
 * @param width Width of the frame
 * @param height Height of the frame
 * @param channels Number of channels (usually 3 for RGB)
 * @param result Pointer to store the detection result
 * @param stream_name Name of the stream (needed for motion detection)
 * @param frame_time Timestamp of the frame (needed for motion detection)
 * @return 0 on success, -1 on failure
 */
int detect_objects(detection_model_t model, const unsigned char *frame_data,
                  int width, int height, int channels, detection_result_t *result,
                  const char *stream_name, time_t frame_time);

#endif /* DETECTION_H */
