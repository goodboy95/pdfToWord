using OpenCvSharp;
using OpenCvSharp.Extensions;
using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;

namespace Pdf2Word.Infrastructure.ImageProcessing;

public sealed class OpenCvPageImagePipeline : IPageImagePipeline
{
    public PageImageBundle Process(System.Drawing.Bitmap renderedPage, CropOptions crop, PreprocessOptions options, int pageNumber)
    {
        var originalSize = (renderedPage.Width, renderedPage.Height);
        var cropped = ApplyCrop(renderedPage, crop, out var cropInfo);

        using var croppedMat = BitmapConverter.ToMat(cropped);
        using var gray = new Mat();
        Cv2.CvtColor(croppedMat, gray, ColorConversionCodes.BGR2GRAY);

        using var enhanced = ApplyContrast(gray, options);
        using var denoised = ApplyDenoise(enhanced, options);
        using var binary = ApplyBinarize(denoised, options);

        var angle = options.EnableDeskew ? ComputeDeskewAngle(binary, options) : 0.0;
        using var deskewed = options.EnableDeskew && Math.Abs(angle) > 0 ? Rotate(binary, angle) : binary.Clone();
        using var colorForGemini = options.EnableDeskew && Math.Abs(angle) > 0 ? Rotate(croppedMat, angle) : croppedMat.Clone();

        var grayBitmap = enhanced.ToBitmap();
        var binaryBitmap = deskewed.ToBitmap();
        var colorBitmap = colorForGemini.ToBitmap();

        var bundle = new PageImageBundle
        {
            PageNumber = pageNumber,
            OriginalColor = renderedPage,
            CroppedColor = cropped,
            Gray = grayBitmap,
            Binary = binaryBitmap,
            ColorForGemini = colorBitmap,
            BinaryForTable = binaryBitmap,
            CropInfo = cropInfo,
            OriginalSizePx = originalSize,
            CroppedSizePx = (cropped.Width, cropped.Height)
        };

        return bundle;
    }

    private static System.Drawing.Bitmap ApplyCrop(System.Drawing.Bitmap source, CropOptions crop, out CropInfo cropInfo)
    {
        var top = 0;
        var bottom = 0;
        if (crop.Mode == HeaderFooterRemoveMode.RemoveHeader || crop.Mode == HeaderFooterRemoveMode.RemoveBoth)
        {
            top = (int)Math.Round(source.Height * crop.HeaderPercent);
        }

        if (crop.Mode == HeaderFooterRemoveMode.RemoveFooter || crop.Mode == HeaderFooterRemoveMode.RemoveBoth)
        {
            bottom = (int)Math.Round(source.Height * crop.FooterPercent);
        }

        var height = Math.Max(1, source.Height - top - bottom);
        var rect = new System.Drawing.Rectangle(0, Math.Max(0, top), source.Width, height);
        var cropped = source.Clone(rect, source.PixelFormat);

        cropInfo = new CropInfo
        {
            Mode = crop.Mode,
            CropTopPx = top,
            CropBottomPx = bottom
        };

        return cropped;
    }

    private static Mat ApplyContrast(Mat gray, PreprocessOptions options)
    {
        return options.ContrastEnhance switch
        {
            ContrastEnhanceMode.None => gray.Clone(),
            ContrastEnhanceMode.Linear => ApplyLinearStretch(gray),
            _ => ApplyClahe(gray, options.Clahe)
        };
    }

    private static Mat ApplyLinearStretch(Mat gray)
    {
        var result = new Mat();
        Cv2.Normalize(gray, result, 0, 255, NormTypes.MinMax);
        return result;
    }

    private static Mat ApplyClahe(Mat gray, ClaheOptions options)
    {
        using var clahe = Cv2.CreateCLAHE(options.ClipLimit, new Size(options.TileGridSize, options.TileGridSize));
        var result = new Mat();
        clahe.Apply(gray, result);
        return result;
    }

    private static Mat ApplyDenoise(Mat gray, PreprocessOptions options)
    {
        var result = new Mat();
        switch (options.Denoise)
        {
            case DenoiseMode.Gaussian3:
                Cv2.GaussianBlur(gray, result, new Size(3, 3), 0);
                break;
            case DenoiseMode.Median3:
                Cv2.MedianBlur(gray, result, 3);
                break;
            default:
                return gray.Clone();
        }

        return result;
    }

    private static Mat ApplyBinarize(Mat gray, PreprocessOptions options)
    {
        var result = new Mat();
        switch (options.Binarize)
        {
            case BinarizeMode.Otsu:
                Cv2.Threshold(gray, result, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
                break;
            default:
                var blockSize = options.Adaptive.BlockSize % 2 == 0 ? options.Adaptive.BlockSize + 1 : options.Adaptive.BlockSize;
                blockSize = Math.Max(3, blockSize);
                Cv2.AdaptiveThreshold(gray, result, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, blockSize, options.Adaptive.C);
                break;
        }

        return result;
    }

    private static double ComputeDeskewAngle(Mat binary, PreprocessOptions options)
    {
        using var edges = new Mat();
        Cv2.Canny(binary, edges, 50, 150);
        var lines = Cv2.HoughLinesP(edges, 1, Math.PI / 180, 100, minLineLength: binary.Width / 4.0, maxLineGap: 10);
        if (lines.Length == 0)
        {
            return 0.0;
        }

        var angles = new List<double>();
        foreach (var line in lines)
        {
            var angle = Math.Atan2(line.P2.Y - line.P1.Y, line.P2.X - line.P1.X) * 180.0 / Math.PI;
            if (Math.Abs(angle) < 45)
            {
                angles.Add(angle);
            }
        }

        if (angles.Count == 0)
        {
            return 0.0;
        }

        var median = angles.OrderBy(a => a).ElementAt(angles.Count / 2);
        if (Math.Abs(median) < options.Deskew.MinAngleDeg || Math.Abs(median) > options.Deskew.MaxAngleDeg)
        {
            return 0.0;
        }

        return median;
    }

    private static Mat Rotate(Mat source, double angle)
    {
        var center = new Point2f(source.Width / 2f, source.Height / 2f);
        var rotation = Cv2.GetRotationMatrix2D(center, angle, 1.0);
        var rotated = new Mat();
        Cv2.WarpAffine(source, rotated, rotation, source.Size(), InterpolationFlags.Linear, BorderTypes.Constant, Scalar.White);
        return rotated;
    }
}
