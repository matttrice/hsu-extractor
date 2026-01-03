#!/bin/bash

# extract_pptx_images.sh
# Extracts all embedded images from a .pptx file.
# Output folder is always created next to this script.

set -euo pipefail  # Better error handling

# Get the directory where this script is located
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPT_NAME="$(basename "$0")"

# Usage help
if [[ $# -lt 1 ]] || [[ "$1" == "-h" ]] || [[ "$1" == "--help" ]]; then
    echo "Usage: $SCRIPT_NAME <presentation.pptx> [custom_output_folder_name]"
    echo ""
    echo "Examples:"
    echo "  $SCRIPT_NAME my_slides.pptx"
    echo "  $SCRIPT_NAME presentation.pptx my_images"
    echo ""
    echo "The output folder will always be created next to this script."
    exit 1
fi

PPTX_FILE="$1"
CUSTOM_OUTPUT="${2:-}"  # Optional second argument

# Check if the PPTX file exists
if [[ ! -f "$PPTX_FILE" ]]; then
    echo "Error: File not found: $PPTX_FILE"
    exit 1
fi

# Determine output folder name
if [[ -n "$CUSTOM_OUTPUT" ]]; then
    OUTPUT_DIR="$SCRIPT_DIR/$CUSTOM_OUTPUT"
else
    BASE_NAME="$(basename "$PPTX_FILE" .pptx)"
    OUTPUT_DIR="$SCRIPT_DIR/extracted_images_$BASE_NAME"
fi

# Create output directory (next to the script)
mkdir -p "$OUTPUT_DIR"

# Create a temporary directory for unzipping
TEMP_DIR=$(mktemp -d)
trap 'rm -rf "$TEMP_DIR"' EXIT  # Auto-cleanup on exit or error

echo "Extracting images from: $PPTX_FILE"
echo "Output folder: $OUTPUT_DIR"

# Unzip the PPTX quietly into the temp directory
unzip -q "$PPTX_FILE" -d "$TEMP_DIR"

# Copy all media files (images) to the output folder
if [[ -d "$TEMP_DIR/ppt/media" ]]; then
    cp "$TEMP_DIR/ppt/media/"* "$OUTPUT_DIR"/ 2>/dev/null || true
    IMAGE_COUNT=$(ls -1 "$OUTPUT_DIR" 2>/dev/null | wc -l)
    if [[ $IMAGE_COUNT -eq 0 ]]; then
        echo "Warning: No images found in the presentation."
        rmdir "$OUTPUT_DIR" 2>/dev/null || true  # Remove empty folder
    else
        echo "Success: $IMAGE_COUNT image(s) extracted to:"
        echo "    $OUTPUT_DIR"
    fi
else
    echo "Warning: No media folder found — this presentation may contain no embedded images."
    rmdir "$OUTPUT_DIR" 2>/dev/null || true
fi