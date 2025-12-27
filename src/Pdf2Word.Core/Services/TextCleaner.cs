namespace Pdf2Word.Core.Services;

public static class TextCleaner
{
    public static string RemoveControlChars(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return string.Empty;
        }

        var buffer = new char[input.Length];
        var idx = 0;
        foreach (var ch in input)
        {
            if (ch == '\n' || ch == '\t' || ch == '\r')
            {
                buffer[idx++] = ch;
                continue;
            }

            if (char.IsControl(ch))
            {
                continue;
            }

            buffer[idx++] = ch;
        }

        return new string(buffer, 0, idx);
    }

    public static double ReplacementCharRate(string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return 0;
        }

        var count = input.Count(ch => ch == '\uFFFD');
        return (double)count / input.Length;
    }
}
