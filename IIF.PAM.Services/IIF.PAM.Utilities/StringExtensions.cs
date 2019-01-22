using System;

namespace IIF.PAM.Utilities
{
    public static class StringExtensions
    {
        public static string Left(this string value, int length)
        {
            string result = value;
            if (value == null)
            {
                return null;
            }
            else
            {
                if (length < 0)
                {
                    throw new ArgumentOutOfRangeException("Length cannot be less than or equal to zero.");
                }
                if (result.Length > length)
                {
                    result = result.Substring(0, length);
                }
            }
            return result;
        }

        public static string Right(this string value, int length)
        {
            string result = value;
            if (value == null)
            {
                return null;
            }
            else
            {
                if (length < 0)
                {
                    throw new ArgumentOutOfRangeException("Length cannot be less than or equal to zero.");
                }
                if (result.Length > length)
                {
                    result = result.Substring(result.Length - length, length);
                }
            }
            return result;
        }


        public static string AppendPath(this string value, string pathSeparator, string additionalPath)
        {
            if (value == null)
            {
                throw new ArgumentNullException("value");
            }
            if (pathSeparator == null)
            {
                throw new ArgumentNullException("pathSeparator");
            }
            if (additionalPath == null)
            {
                throw new ArgumentNullException("additionalPath");
            }

            string firstPart = value;
            string lastPart = additionalPath;

            if (firstPart.Right(pathSeparator.Length) == pathSeparator)
            {
                firstPart = firstPart.Left(firstPart.Length - pathSeparator.Length);
            }

            if (lastPart.Left(pathSeparator.Length) == pathSeparator)
            {
                lastPart = lastPart.Right(lastPart.Length - pathSeparator.Length);
            }

            return firstPart + pathSeparator + lastPart;
        }

        public static string AppendFolderPath(this string value, string additionalPath)
        {
            return value.AppendPath("\\", additionalPath);
        }

        public static string AppendUrlPath(this string value, string additionalPath)
        {
            return value.AppendPath("/", additionalPath);
        }
    }
}
