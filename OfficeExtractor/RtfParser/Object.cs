using System;

namespace OfficeExtractor.RtfParser
{
    public class Object
    {
        public string Text { get; private set; }

        public Object(string text)
        {
            if (text == null)
                throw new ArgumentNullException("text");

            Text = text.Trim();
        }
    }
}