using System;
using System.Collections.Generic;
using System.Text;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// interface for compression adapter
    /// </summary>
    /// <author>linwei</author>
    public interface IUofCompressAdapter
    {
        /// <summary>
        /// compress or decompress
        /// </summary>
        /// <returns>true if succeeded</returns>
        /// <returns>false if failed</returns>
        /// <exception cref="Exception">exceptions happen</exception>
        bool Transform();

        /// <summary>
        /// input file name
        /// </summary>
        string InputFilename
        {
            get;
            set;
        }

        /// <summary>
        /// output file name
        /// </summary>
        string OutputFilename
        {
            get;
            set;
        }

    }
}
