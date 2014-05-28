using System;

namespace DocumentServices.Modules.Extractors.OfficeExtractor.CompoundFileStorage.BinaryTree.Exceptions
{
    public class BSTDuplicatedException : ApplicationException
    {
        public BSTDuplicatedException() : base("Duplicated item already present in BSTree") { }
    }
}
