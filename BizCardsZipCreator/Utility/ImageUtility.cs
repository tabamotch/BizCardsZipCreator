using System.Drawing;

namespace jp.tabamotch.BizCardsZipCreator.Utility
{
    public class ImageUtility
    {
        public static Image ConvertImageSize(Image source)
        {
            decimal originalHeight = source.Height;
            decimal originalWidth = source.Width;

            decimal newWidth = originalWidth * (20m / originalHeight);

            Bitmap canvas = new Bitmap(source, (int)newWidth, 20);

            return canvas;
        }
    }
}
