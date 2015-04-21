using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

public class RtfReader
{
    public RtfReader(TextReader reader)
    {
        if (reader == null)
            throw new ArgumentNullException("reader");

        Reader = reader;
    }

    public TextReader Reader { get; private set; }

    public IEnumerable<RtfObject> Read()
    {
        StringBuilder controlWord = new StringBuilder();
        StringBuilder text = new StringBuilder();
        Stack<RtfParseState> stack = new Stack<RtfParseState>();
        RtfParseState state = RtfParseState.Group;

        do
        {
            int i = Reader.Read();
            if (i < 0)
            {
                if (!string.IsNullOrWhiteSpace(controlWord.ToString()))
                    yield return new RtfControlWord(controlWord.ToString());

                if (!string.IsNullOrWhiteSpace(text.ToString()))
                    yield return new RtfText(text.ToString());

                yield break;
            }

            char c = (char)i;

            // noise chars
            if ((c == '\r') ||
                (c == '\n'))
                continue;

            switch (state)
            {
                case RtfParseState.Group:
                    if (c == '{')
                    {
                        stack.Push(state);
                        break;
                    }

                    if (c == '\\')
                    {
                        state = RtfParseState.ControlWord;
                        break;
                    }
                    break;

                case RtfParseState.ControlWord:
                    if (c == '\\')
                    {
                        // another controlWord
                        if (!string.IsNullOrWhiteSpace(controlWord.ToString()))
                        {
                            yield return new RtfControlWord(controlWord.ToString());
                            controlWord.Clear();
                        }
                        break;
                    }

                    if (c == '{')
                    {
                        // a new group
                        state = RtfParseState.Group;
                        if (!string.IsNullOrWhiteSpace(controlWord.ToString()))
                        {
                            yield return new RtfControlWord(controlWord.ToString());
                            controlWord.Clear();
                        }
                        break;
                    }

                    if (c == '}')
                    {
                        // close group
                        state = stack.Count > 0 ? stack.Pop() : RtfParseState.Group;
                        if (!string.IsNullOrWhiteSpace(controlWord.ToString()))
                        {
                            yield return new RtfControlWord(controlWord.ToString());
                            controlWord.Clear();
                        }
                        break;
                    }

                    if (!Char.IsLetterOrDigit(c))
                    {
                        state = RtfParseState.Text;
                        text.Append(c);
                        if (!string.IsNullOrWhiteSpace(controlWord.ToString()))
                        {
                            yield return new RtfControlWord(controlWord.ToString());
                            controlWord.Clear();
                        }
                        break;
                    }

                    controlWord.Append(c);
                    break;

                case RtfParseState.Text:
                    if (c == '\\')
                    {
                        state = RtfParseState.EscapedText;
                        break;
                    }

                    if (c == '{')
                    {
                        if (!string.IsNullOrWhiteSpace(text.ToString()))
                        {
                            yield return new RtfText(text.ToString());
                            text.Clear();
                        }

                        // a new group
                        state = RtfParseState.Group;
                        break;
                    }

                    if (c == '}')
                    {
                        if (!string.IsNullOrWhiteSpace(text.ToString()))
                        {
                            yield return new RtfText(text.ToString());
                            text.Clear();
                        }

                        // close group
                        state = stack.Count > 0 ? stack.Pop() : RtfParseState.Group;
                        break;
                    }
                    text.Append(c);
                    break;

                case RtfParseState.EscapedText:
                    if ((c == '\\') || (c == '}') || (c == '{'))
                    {
                        state = RtfParseState.Text;
                        text.Append(c);
                        break;
                    }

                    // ansi character escape
                    if (c == '\'')
                    {
                        text.Append(FromHexa((char)Reader.Read(), (char)Reader.Read()));
                        break;
                    }

                    if (!string.IsNullOrWhiteSpace(text.ToString()))
                    {
                        yield return new RtfText(text.ToString());
                        text.Clear();
                    }

                    // in fact, it's a normal controlWord
                    controlWord.Append(c);
                    state = RtfParseState.ControlWord;
                    break;
            }
        }
        while (true);
    }

    public static bool MoveToNextControlWord(IEnumerator<RtfObject> enumerator, string word)
    {
        if (enumerator == null)
            throw new ArgumentNullException("enumerator");

        while (enumerator.MoveNext())
        {
            if (enumerator.Current.Text == word)
                return true;
        }
        return false;
    }

    public static string GetNextText(IEnumerator<RtfObject> enumerator)
    {
        if (enumerator == null)
            throw new ArgumentNullException("enumerator");

        while (enumerator.MoveNext())
        {
            RtfText text = enumerator.Current as RtfText;
            if (text != null)
                return text.Text;
        }
        return null;
    }

    public static byte[] GetNextTextAsByteArray(IEnumerator<RtfObject> enumerator)
    {
        if (enumerator == null)
            throw new ArgumentNullException("enumerator");

        while (enumerator.MoveNext())
        {
            RtfText text = enumerator.Current as RtfText;
            if (text != null)
            {
                List<byte> bytes = new List<byte>();
                for (int i = 0; i < text.Text.Length; i += 2)
                {
                    bytes.Add((byte)FromHexa(text.Text[i], text.Text[i + 1]));
                }
                return bytes.ToArray();
            }
        }
        return null;
    }

    // Extracts an EmbeddedObject/ObjectHeader from a stream
    // see [MS -OLEDS]: Object Linking and Embedding (OLE) Data Structures for more information
    // chapter 2.2: OLE1.0 Format Structures 
    public static void ExtractObjectData(Stream inputStream, Stream outputStream)
    {
        if (inputStream == null)
            throw new ArgumentNullException("inputStream");

        if (outputStream == null)
            throw new ArgumentNullException("outputStream");

        BinaryReader reader = new BinaryReader(inputStream);
        reader.ReadInt32(); // OLEVersion
        int formatId = reader.ReadInt32(); // FormatID
        if (formatId != 2) // see 2.2.4 Object Header. 2 means EmbeddedObject
            throw new NotSupportedException();

        ReadLengthPrefixedAnsiString(reader); // className
        ReadLengthPrefixedAnsiString(reader); // topicName
        ReadLengthPrefixedAnsiString(reader); // itemName

        int nativeDataSize = reader.ReadInt32();
        byte[] bytes = reader.ReadBytes(nativeDataSize);
        outputStream.Write(bytes, 0, bytes.Length);
    }

    // see chapter 2.1.4 LengthPrefixedAnsiString
    private static string ReadLengthPrefixedAnsiString(BinaryReader reader)
    {
        int length = reader.ReadInt32();
        if (length == 0)
            return string.Empty;

        byte[] bytes = reader.ReadBytes(length);
        return Encoding.Default.GetString(bytes, 0, length - 1);
    }

    private enum RtfParseState
    {
        ControlWord,
        Text,
        EscapedText,
        Group
    }

    private static char FromHexa(char hi, char lo)
    {
        return (char)byte.Parse(hi.ToString() + lo, NumberStyles.HexNumber);
    }
}

// Utility class to parse an OLE1.0 OLEOBJECT
public class PackagedObject
{
    private PackagedObject()
    {
    }

    public string DisplayName { get; private set; }
    public string IconFilePath { get; private set; }
    public int IconIndex { get; private set; }
    public string FilePath { get; private set; }
    public byte[] Data { get; private set; }

    private static string ReadAnsiString(BinaryReader reader)
    {
        StringBuilder sb = new StringBuilder();
        do
        {
            byte b = reader.ReadByte();
            if (b == 0)
                return sb.ToString();

            sb.Append((char)b);
        }
        while (true);
    }

    public static PackagedObject Extract(Stream inputStream)
    {
        if (inputStream == null)
            throw new ArgumentNullException("inputStream");

        BinaryReader reader = new BinaryReader(inputStream);
        reader.ReadUInt16(); // sig
        PackagedObject po = new PackagedObject();
        po.DisplayName = ReadAnsiString(reader);
        po.IconFilePath = ReadAnsiString(reader);
        po.IconIndex = reader.ReadUInt16();
        int type = reader.ReadUInt16();
        if (type != 3) // 3 is file, 1 is link
            throw new NotSupportedException();

        reader.ReadInt32(); // nextsize
        po.FilePath = ReadAnsiString(reader);
        int dataSize = reader.ReadInt32();
        po.Data = reader.ReadBytes(dataSize);
        // note after that, there may be unicode + long path info
        return po;
    }
}

public class RtfObject
{
    public RtfObject(string text)
    {
        if (text == null)
            throw new ArgumentNullException("text");

        Text = text.Trim();
    }

    public string Text { get; private set; }
}

public class RtfText : RtfObject
{
    public RtfText(string text)
        : base(text)
    {
    }
}

public class RtfControlWord : RtfObject
{
    public RtfControlWord(string name)
        : base(name)
    {
    }
}