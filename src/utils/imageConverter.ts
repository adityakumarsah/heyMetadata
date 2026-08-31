import { encode } from '@jsquash/avif';

export type SupportedFormat = 'jpg' | 'jpeg' | 'png' | 'webp' | 'avif';

export interface ConversionOptions {
  targetFormat: SupportedFormat;
  quality?: number; // 0.1 to 1.0 (defaults: 0.92 for jpg/jpeg/webp/avif, 1.0 for png)
}

export interface ImageMetadataInfo {
  name: string;
  size: number;
  width: number;
  height: number;
  format: SupportedFormat;
  mimeType: string;
}

export interface ConversionResult {
  blob: Blob;
  dataUrl: string;
  fileName: string;
  fileSize: number;
  width: number;
  height: number;
  sourceFormat: SupportedFormat;
  targetFormat: SupportedFormat;
  mimeType: string;
}

export const FORMAT_MIME_MAP: Record<SupportedFormat, string> = {
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  png: 'image/png',
  webp: 'image/webp',
  avif: 'image/avif',
};

export const EXTENSION_FORMAT_MAP: Record<string, SupportedFormat> = {
  jpg: 'jpg',
  jpeg: 'jpeg',
  png: 'png',
  webp: 'webp',
  avif: 'avif',
};

/**
 * Valid conversions mapping:
 * - JPG -> JPEG, PNG, WEBP, AVIF
 * - JPEG -> JPG, PNG, WEBP, AVIF
 * - PNG -> JPG, JPEG, WEBP, AVIF
 * - WEBP -> JPG, JPEG, PNG, AVIF
 * - AVIF -> JPG, JPEG, PNG, WEBP
 */
export const ALLOWED_CONVERSIONS: Record<SupportedFormat, SupportedFormat[]> = {
  jpg: ['jpeg', 'png', 'webp', 'avif'],
  jpeg: ['jpg', 'png', 'webp', 'avif'],
  png: ['jpg', 'jpeg', 'webp', 'avif'],
  webp: ['jpg', 'jpeg', 'png', 'avif'],
  avif: ['jpg', 'jpeg', 'png', 'webp'],
};

/**
 * Cache for browser native AVIF encoding capability test.
 */
let _nativeAvifSupported: boolean | null = null;

/**
 * Check if the current browser environment supports native encoding to image/avif via canvas.toBlob.
 */
export async function isNativeAvifEncodeSupported(): Promise<boolean> {
  if (typeof document === 'undefined') return false;
  if (_nativeAvifSupported !== null) return _nativeAvifSupported;

  try {
    const testCanvas = document.createElement('canvas');
    testCanvas.width = 1;
    testCanvas.height = 1;
    const blob = await new Promise<Blob | null>((resolve) => {
      testCanvas.toBlob((b) => resolve(b), 'image/avif', 0.8);
    });

    // Standard HTML Canvas specification states toBlob falls back to image/png if the format is unsupported
    _nativeAvifSupported = blob !== null && blob.type === 'image/avif';
  } catch {
    _nativeAvifSupported = false;
  }

  return _nativeAvifSupported;
}

/**
 * Encodes canvas image data to AVIF using client-side WASM (libavif via @jsquash/avif).
 */
async function encodeAvifWithWasm(
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number,
  quality: number
): Promise<Blob> {
  try {
    const imageData = ctx.getImageData(0, 0, width, height);
    const qualityInt = Math.min(Math.max(Math.round(quality * 100), 1), 100);
    const buffer = await encode(imageData, { quality: qualityInt });
    if (!buffer || buffer.byteLength === 0) {
      throw new Error('AVIF WASM encoder produced an empty output buffer.');
    }
    return new Blob([buffer], { type: 'image/avif' });
  } catch (err: any) {
    throw new Error(`Client-side AVIF encoding error: ${err?.message || 'Unknown error'}`);
  }
}

/**
 * Check if format is a JPEG variant (jpg or jpeg).
 */
export function isJpegVariant(format: SupportedFormat): boolean {
  return format === 'jpg' || format === 'jpeg';
}

/**
 * Detect the image format from File name extension or MIME type.
 * Distinguishes between .jpg and .jpeg file extensions.
 */
export function detectFormat(file: { name?: string; type?: string }): SupportedFormat | null {
  const ext = (file.name ? file.name.split('.').pop() || '' : '').toLowerCase();
  if (EXTENSION_FORMAT_MAP[ext]) {
    return EXTENSION_FORMAT_MAP[ext];
  }

  if (file.type) {
    if (file.type === 'image/jpeg' || file.type === 'image/jpg') {
      return 'jpg';
    }
    if (file.type === 'image/png') return 'png';
    if (file.type === 'image/webp') return 'webp';
    if (file.type === 'image/avif') return 'avif';
  }

  return null;
}

/**
 * Check if a conversion from sourceFormat to targetFormat is supported.
 */
export function isConversionSupported(sourceFormat: SupportedFormat, targetFormat: SupportedFormat): boolean {
  if (sourceFormat === targetFormat) return false;
  const targets = ALLOWED_CONVERSIONS[sourceFormat];
  return targets ? targets.includes(targetFormat) : false;
}

/**
 * Get the list of allowed target formats for a given source format.
 */
export function getAvailableTargetFormats(sourceFormat: SupportedFormat): SupportedFormat[] {
  return ALLOWED_CONVERSIONS[sourceFormat] || [];
}

/**
 * Formats byte size into human readable string.
 */
export function formatBytes(bytes: number): string {
  if (bytes === 0) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${Math.ceil(bytes / 1024)} KB`;
  return `${(bytes / 1048576).toFixed(2)} MB`;
}

/**
 * Validates and inspects an image file entirely in-browser.
 */
export async function inspectImageFile(file: File): Promise<ImageMetadataInfo> {
  if (!file) {
    throw new Error('No file provided.');
  }

  if (file.size === 0) {
    throw new Error('The selected file is empty (0 bytes).');
  }

  const format = detectFormat(file);
  if (!format) {
    throw new Error('Unsupported image format. Please select a JPG, JPEG, PNG, WEBP, or AVIF image.');
  }

  // Attempt to decode image to inspect dimensions and verify file is not corrupted
  let bitmap: ImageBitmap;
  try {
    bitmap = await createImageBitmap(file);
  } catch {
    throw new Error(`Unable to read image. The file may be invalid, corrupted, or ${format.toUpperCase()} decoding is unsupported in your browser.`);
  }

  const width = bitmap.width;
  const height = bitmap.height;
  bitmap.close();

  if (width === 0 || height === 0) {
    throw new Error('Image has invalid dimensions.');
  }

  return {
    name: file.name,
    size: file.size,
    width,
    height,
    format,
    mimeType: FORMAT_MIME_MAP[format],
  };
}

/**
 * Generates an output filename with the exact new extension.
 */
export function generateOutputFileName(originalName: string, targetFormat: SupportedFormat): string {
  const baseName = originalName.replace(/\.[^/.]+$/, '');
  return `${baseName}.${targetFormat}`;
}

/**
 * Converts an image file client-side using Canvas / ImageBitmap / WASM APIs.
 */
export async function convertImage(
  file: File,
  options: ConversionOptions
): Promise<ConversionResult> {
  if (!file || file.size === 0) {
    throw new Error('Invalid file for conversion.');
  }

  const sourceFormat = detectFormat(file);
  if (!sourceFormat) {
    throw new Error('Unsupported source image format. Only JPG, JPEG, PNG, WEBP, and AVIF are supported.');
  }

  const { targetFormat, quality = 0.92 } = options;

  if (!isConversionSupported(sourceFormat, targetFormat)) {
    throw new Error(`Conversion from ${sourceFormat.toUpperCase()} to ${targetFormat.toUpperCase()} is not supported.`);
  }

  let bitmap: ImageBitmap;
  try {
    bitmap = await createImageBitmap(file);
  } catch {
    throw new Error('Failed to decode image data. The file may be corrupted or unsupported by your browser.');
  }

  const width = bitmap.width;
  const height = bitmap.height;

  const isTargetJpeg = isJpegVariant(targetFormat);

  // Create canvas with matching dimensions
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;

  const ctx = canvas.getContext('2d', { alpha: !isTargetJpeg });
  if (!ctx) {
    bitmap.close();
    throw new Error('Browser 2D canvas context is not available.');
  }

  // If converting to JPEG/JPG, fill canvas with white background so transparent areas don't render black
  if (isTargetJpeg) {
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, width, height);
  }

  // Draw bitmap onto canvas preserving 1:1 dimensions and pixels
  ctx.drawImage(bitmap, 0, 0, width, height);
  bitmap.close();

  const targetMime = FORMAT_MIME_MAP[targetFormat];
  const targetQuality = targetFormat === 'png' ? 1.0 : Math.min(Math.max(quality, 0.1), 1.0);

  let blob: Blob;

  if (targetFormat === 'avif') {
    // 1. Try native canvas AVIF encoding if supported
    const nativeSupported = await isNativeAvifEncodeSupported();
    let nativeSuccess = false;

    if (nativeSupported) {
      try {
        blob = await new Promise<Blob>((resolve, reject) => {
          canvas.toBlob(
            (b) => {
              if (b && b.type === 'image/avif') {
                resolve(b);
              } else {
                reject(new Error('Native AVIF encoder did not return image/avif'));
              }
            },
            'image/avif',
            targetQuality
          );
        });
        nativeSuccess = true;
      } catch {
        nativeSuccess = false;
      }
    }

    // 2. Fall back to client-side WASM AVIF encoder
    if (!nativeSuccess) {
      blob = await encodeAvifWithWasm(ctx, width, height, targetQuality);
    }
  } else {
    // Standard canvas encoding for JPG, PNG, WEBP
    blob = await new Promise<Blob>((resolve, reject) => {
      canvas.toBlob(
        (b) => {
          if (!b) {
            reject(new Error(`Browser failed to encode image into ${targetFormat.toUpperCase()} format.`));
            return;
          }
          resolve(b);
        },
        targetMime,
        targetQuality
      );
    });
  }

  const fileName = generateOutputFileName(file.name, targetFormat);
  const dataUrl = URL.createObjectURL(blob);

  return {
    blob,
    dataUrl,
    fileName,
    fileSize: blob.size,
    width,
    height,
    sourceFormat,
    targetFormat,
    mimeType: targetMime,
  };
}
