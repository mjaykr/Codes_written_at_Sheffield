# Animated GIF Creator from TIFF Sequence

This MATLAB script automates the process of creating an animated GIF from a sequence of TIFF images. It reads all the `.tif` files present in the directory where the script is located, resizes the images to a width of 1080 pixels while maintaining the aspect ratio, and compiles them into an optimized animated GIF that loops indefinitely.

## Prerequisites

- MATLAB (The script has been tested on MATLAB R2022a and later versions)
- Image Processing Toolbox (For image reading and writing functions)

## Usage

1. Place the MATLAB script in the directory containing your `.tif` images.
2. Ensure that the `.tif` images are named in a sequence if you want them to appear in a specific order.
3. Run the script in MATLAB.

The script will create an animated GIF named `animated_sequence.gif` in the current directory.

## Features

- **Automatic TIFF Image Detection**: The script automatically detects and processes all `.tif` files in the directory.
- **Aspect Ratio Maintenance**: Resizes images to a fixed width of 1080 pixels, maintaining the original aspect ratio to avoid distortion.
- **Grayscale and Indexed Image Compatibility**: Converts grayscale and indexed images to RGB before processing, ensuring compatibility.
- **Bit Depth Handling**: Ensures images are converted to 8-bit before indexing, which is a requirement for GIF creation.
- **Optimized Frame Timing**: Sets a default frame delay time of 0.1 seconds between frames, which can be adjusted as needed.
- **Infinite Looping**: The generated GIF will loop indefinitely, making it suitable for presentations or web usage.

## Customization

The script can be easily customized:
- The `targetWidth` variable can be changed to resize the images to a different width.
- The `DelayTime` parameter in the `imwrite` function can be adjusted to speed up or slow down the animation.

## Troubleshooting

Ensure that all TIFF files are in an uncompressed, non-layered format. The script does not support TIFF images with layers or special compression schemes that MATLAB cannot read directly.

## Author

This script was developed by [Dr Mirtunjay Kumar](https://mjaykr.github.io/).

For any issues or bugs, please open an issue on the GitHub repository or contact [mjay@hotmail.com](mailto:mjay@hotmail.com).

## License

This project is open-sourced under the MIT License. See the LICENSE file for more details.
