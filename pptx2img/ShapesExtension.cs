using Microsoft.Office.Interop.PowerPoint;

namespace pptx2img
{
    public static class ShapesExtension
    {
        public static int[] GetIndices(this Shapes shapes)
        {
            int[] result = new int[shapes.Count];
            for (int i = 1; i <= shapes.Count; i++)
            {
                result[i - 1] = i;
            }

            return result;
        }
    }
}