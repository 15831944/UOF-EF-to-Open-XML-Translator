/*
 * Copyright (c) 2006, Tsinghua University, China
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age, nor the names of its contributors may
 *       be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace UofAddinLib
{
    /// <summary>
    /// Module ID:      
    /// Depiction:      获取文件系统中对象的信息，例如：文件、文件夹、驱动器根目录 
    /// Author:         http://www.cnblogs.com/BG5SBK/articles/GetFileInfo.aspx
    /// Create Date:    2007-07-01
    /// </summary>   
    public class FileInfo
    {
        [DllImport("shell32.dll", EntryPoint = "SHGetFileInfo", CharSet = CharSet.Auto)]
        public static extern int GetFileInfo(string pszPath, int dwFileAttributes,
         ref  FileInfomation psfi, int cbFileInfo, int uFlags);

        private FileInfo() { }

        /// <summary>
        ///     Make A File Name Based on the Original File Name
        /// </summary>
        /// <param name="baseFileName">
        ///     The Original File Name. For example, it can be "Test"
        ///     And the return will be like "Test.xxx"
        /// </param>
        /// <param name="extendFileName">
        ///     The Extend File Name. For example, it can be ".docx"
        ///     And the return will be like "xxx.docx"
        /// </param>
        /// <returns>
        ///     The File Name That is Made.
        /// </returns>
        ///     Author:         邓追
        ///     Create Date:    2007-07-4
        public static string MakeFileName(string baseFileName, string extendFileName)
        {
            string outputFileName = baseFileName;
            string extFileName = extendFileName;
            UInt64 count = 1;
            bool isOK = false;

            if (!System.IO.File.Exists(outputFileName + extFileName))
            {
                outputFileName = outputFileName + extFileName;
                return outputFileName;
            }

            while (!isOK)
            {
                isOK = true;
                if (System.IO.File.Exists(outputFileName + "(" + count.ToString() + ")" + extFileName))
                {
                    // If the file already exist,test if it can be opened exclusively
                    count++;
                    isOK = false;
                }
            }

            outputFileName = outputFileName + "(" + count.ToString() + ")" + extFileName;
            return outputFileName;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct FileInfomation
    {
        public IntPtr hIcon;
        public int iIcon;
        public int dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    public enum FileAttributeFlags : int
    {
        FILE_ATTRIBUTE_READONLY = 0x00000001,
        FILE_ATTRIBUTE_HIDDEN = 0x00000002,
        FILE_ATTRIBUTE_SYSTEM = 0x00000004,
        FILE_ATTRIBUTE_DIRECTORY = 0x00000010,
        FILE_ATTRIBUTE_ARCHIVE = 0x00000020,
        FILE_ATTRIBUTE_DEVICE = 0x00000040,
        FILE_ATTRIBUTE_NORMAL = 0x00000080,
        FILE_ATTRIBUTE_TEMPORARY = 0x00000100,
        FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200,
        FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400,
        FILE_ATTRIBUTE_COMPRESSED = 0x00000800,
        FILE_ATTRIBUTE_OFFLINE = 0x00001000,
        FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000,
        FILE_ATTRIBUTE_ENCRYPTED = 0x00004000
    }

    public enum GetFileInfoFlags : int
    {
        SHGFI_ICON = 0x000000100,               //  get icon 
        SHGFI_DISPLAYNAME = 0x000000200,        //  get display name 
        SHGFI_TYPENAME = 0x000000400,           //  get type name 
        SHGFI_ATTRIBUTES = 0x000000800,         //  get attributes 
        SHGFI_ICONLOCATION = 0x000001000,       //  get icon location 
        SHGFI_EXETYPE = 0x000002000,            //  return exe type 
        SHGFI_SYSICONINDEX = 0x000004000,       //  get system icon index 
        SHGFI_LINKOVERLAY = 0x000008000,        //  put a link overlay on icon 
        SHGFI_SELECTED = 0x000010000,           //  show icon in selected state 
        SHGFI_ATTR_SPECIFIED = 0x000020000,     //  get only specified attributes 
        SHGFI_LARGEICON = 0x000000000,          //  get large icon 
        SHGFI_SMALLICON = 0x000000001,          //  get small icon 
        SHGFI_OPENICON = 0x000000002,           //  get open icon 
        SHGFI_SHELLICONSIZE = 0x000000004,      //  get shell size icon 
        SHGFI_PIDL = 0x000000008,               //  pszPath is a pidl 
        SHGFI_USEFILEATTRIBUTES = 0x000000010,  //  use passed dwFileAttribute 
        SHGFI_ADDOVERLAYS = 0x000000020,        //  apply the appropriate overlays 
        SHGFI_OVERLAYINDEX = 0x000000040        //  Get the index of the overlay 
    }
}
