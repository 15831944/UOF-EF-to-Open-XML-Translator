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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace UofAddinLib
{
    /// <summary>
    ///     Main Dialog Types: For Open File or Save File
    /// </summary>
    public enum DialogType 
    {
        /// <summary>
        ///     This is the Dialog Type used when Open UOF File in Word
        /// </summary>
        Open,

        /// <summary>
        ///     This is the Dialog Type used when Save UOF File in Word
        /// </summary>
        Save,

        /// <summary>
        ///     This is the Dialog Type used when translate UOF File from Explorer
        /// </summary>
        UofToOox,

        /// <summary>
        ///     This is the Dialog Type used when translate Oox File from Explorer
        /// </summary>
        OoxToUof
    };

    /// <summary>
    ///     DocumentType: Use this to decide which Converter to use
    /// </summary>
    public enum DocumentType
    {
        Unknown,

        Word,

        Excel,

        Powerpnt,
        
        StrictWord,

        StrictExcel,

        StrictPowerpnt,

        All,         //仅对整目录生效

        UofPending   //UOF文件但尚未确定具体类型
    };

    public enum TranslationType
    {
        Unknown,
        UofToOox,
        OoxToUof,
        UofToTransitionalOOX,
        UofToStrictOOX
    }

    public static class Extensions
    {
        public const string OOX_WORD = ".docx";
        public const string OOX_WORD_SEARCH = "*.docx";

        public const string OOX_EXCEL = ".xlsx";
        public const string OOX_EXCEL_SEARCH = "*.xlsx";

        public const string OOX_POWERPNT = ".pptx";
        public const string OOX_POWERPNT_SEARCH = "*.pptx";

        //public const string UOF = ".uof";
        //public const string UOF_SEARCH = "*.uof";
        //public const string UOF_SEARCH_PRECISION = "*?.uof";

        public const string UOF_WORD = ".uot";
        public const string UOF_WORD_SEARCH = "*.uot";

        public const string UOF_EXCEL = ".uos";
        public const string UOF_EXCEL_SEARCH = "*.uos";

        public const string UOF_POWERPNT = ".uop";
        public const string UOF_POWERPNT_SEARCH = "*.uop";

        public const string ALL_SEARCH = "*.*";
    }
}