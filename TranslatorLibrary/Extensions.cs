namespace TranslatorLibrary
{
    public static class Extensions
    {
        public static string EnsureStartsWith(this string input, string start)
        {
            if (!input.StartsWith(start))
            {
                return start + input;
            }
            return input;
        }
        public static string EnsureEndsWith(this string input, string end)
        {
            if (!input.EndsWith(end))
            {
                return input + end;
            }
            return input;
        }
        public static string NoSlashes(this string input)
        {
            return input.Replace('/', '.').Replace('\\', '.');
        }
    }
}
