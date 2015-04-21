namespace OfficeExtractor.Biff8.Interfaces
{
    internal interface ILittleEndianOutput
    {
        void WriteByte(int v);
        void WriteShort(int v);
        void WriteInt(int v);
        void WriteLong(long v);
        void WriteDouble(double v);
        void Write(byte[] b);
        void Write(byte[] b, int offset, int len);
    }
}
