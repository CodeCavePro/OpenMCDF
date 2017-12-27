using System;
using System.Diagnostics.CodeAnalysis;

namespace OpenMcdf
{
    /// <summary>
    /// Configuration parameters for the compound files.
    /// They can be OR-combined to configure 
    /// <see cref="T:OpenMcdf.CompoundFile">Compound file</see> behavior.
    /// All flags are NOT set by Default.
    /// </summary>
    [Flags]
    public enum CFSConfiguration
    {
        /// <summary>
        /// Sector Recycling turn off, 
        /// free sectors erasing off, 
        /// format validation exception raised
        /// </summary>
        Default = 1,

        /// <summary>
        /// Sector recycling reduces data writing performances 
        /// but avoids space wasting in scenarios with frequently
        /// data manipulation of the same streams.
        /// </summary>
        SectorRecycle = 2,

        /// <summary>
        /// Free sectors are erased to avoid information leakage
        /// </summary>
        EraseFreeSectors = 4,

        /// <summary>
        /// No exception is raised when a validation error occurs.
        /// This can possibly lead to a security issue but gives 
        /// a chance to corrupted files to load.
        /// </summary>
        NoValidationException = 8,

        /// <summary>
        /// If this flag is set true,
        /// backing stream is kept open after CompoundFile disposal
        /// </summary>
        LeaveOpen = 16,
    }

    /// <summary>
    /// Binary File Format Version. Sector size  is 512 byte for version 3,
    /// 4096 for version 4
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum CFSVersion
    {
        /// <summary>
        /// Compound file version 3 - The default and most common version available. Sector size 512 bytes, 2GB max file size.
        /// </summary>
        Ver_3 = 3,

        /// <summary>
        /// Compound file version 4 - Sector size is 4096 bytes. Using this version could bring some compatibility problem with existing applications.
        /// </summary>
        Ver_4 = 4
    }

    /// <summary>
    /// Update mode of the compound file.
    /// Default is ReadOnly.
    /// </summary>
    public enum CFSUpdateMode
    {
        /// <summary>
        /// ReadOnly update mode prevents overwriting
        /// of the opened file. 
        /// Data changes are allowed but they have to be 
        /// persisted on a different file when required 
        /// using <see cref="M:OpenMcdf.CompoundFile.Save">method</see>
        /// </summary>
        ReadOnly,

        /// <summary>
        /// Update mode allows subsequent data changing operations
        /// to be persisted directly on the opened file or stream
        /// using the <see cref="M:OpenMcdf.CompoundFile.Commit">Commit</see>
        /// method when required. Warning: this option may cause existing data loss if misused.
        /// </summary>
        Update
    }

    public enum StgType
    {
        StgInvalid = 0,
        StgStorage = 1,
        StgStream = 2,
        StgLockbytes = 3,
        StgProperty = 4,
        StgRoot = 5
    }

    public enum StgColor
    {
        Red = 0,
        Black = 1
    }
}
