using System;

namespace OfficeExtractor.Rtf
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