namespace DocumentServices.Modules.Extractors.OfficeExtractor.Biff8.Interfaces
{
    internal interface IBiffHeaderInput
    {
        /// <summary>
        /// Read an unsigned short from the stream without decrypting
        /// </summary>
        /// <returns></returns>
        int ReadRecordSid();
        
        /// <summary>
        /// Read an unsigned short from the stream without decrypting
        /// </summary>
        /// <returns></returns>
        int ReadDataSize();

        int Available();
    }
}
