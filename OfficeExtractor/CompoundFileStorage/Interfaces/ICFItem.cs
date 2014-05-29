using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.Interfaces
{
    /// <summary>
    /// The compound file item interface
    /// </summary>
    public interface ICFItem
    {
        /// <summary>
        ///     Get entity name
        /// </summary>
        string Name { get; }

        /// <summary>
        ///     Size in bytes of the item. It has a valid value
        ///     only if entity is a stream, otherwise it is setted to zero.
        /// </summary>
        long Size { get; }

        /// <summary>
        ///     Return true if item is Storage
        /// </summary>
        /// <remarks>
        ///     This check doesn't use reflection or runtime type information
        ///     and doesn't suffer related performance penalties.
        /// </remarks>
        bool IsStorage { get; }

        /// <summary>
        ///     Return true if item is a Stream
        /// </summary>
        /// <remarks>
        ///     This check doesn't use reflection or runtime type information
        ///     and doesn't suffer related performance penalties.
        /// </remarks>
        bool IsStream { get; }

        /// <summary>
        ///     Return true if item is the Root Storage
        /// </summary>
        /// <remarks>
        ///     This check doesn't use reflection or runtime type information
        ///     and doesn't suffer related performance penalties.
        /// </remarks>
        bool IsRoot { get; }

        /// <summary>
        ///     Get/Set the Creation Date of the current item
        /// </summary>
        DateTime CreationDate { get; set; }

        /// <summary>
        ///     Get/Set the Modify Date of the current item
        /// </summary>
        DateTime ModifyDate { get; set; }

        /// <summary>
        ///     Get/Set Object class Guid for Root and Storage entries.
        /// </summary>
        Guid CLSID { get; set; }
    }
}