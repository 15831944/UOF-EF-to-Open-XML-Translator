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
using System.IO;
using Act.UofTranslator.UofTranslatorLib;

namespace UofAddinLib
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      ÃèÊöÎÄ¼þµÄ×ª»¯×´Ì¬
    ///     Author:         µË×·
    ///     Create Date:    2007-05-25
    /// </summary>
    public class FileState
    {

        private TranslationType implyTransType;

        public TranslationType ImplyTransType
        {
            set
            {
                implyTransType = value;
            }
        }

        private TranslationType transType;

        public TranslationType TransType
        {
            get { return (transType == TranslationType.Unknown ? implyTransType : transType); }
        }

        /// <summary>
        ///     The type of the document
        /// </summary>
        private DocumentType docType;

        public DocumentType DocType
        {
            get{ return (docType == DocumentType.Unknown ? implyDocType : docType); }
        }

        private DocumentType implyDocType;

        public DocumentType ImplyDocType
        {
            set
            {
                if (value == DocumentType.UofPending)
                {
                    try
                    {
                        Act.UofTranslator.UofTranslatorLib.DocType uofDocType = TranslatorFactory.CheckUOFFileType(srcFileName);
                        if (uofDocType == Act.UofTranslator.UofTranslatorLib.DocType.Excel)
                        {
                            implyDocType = DocumentType.Excel;
                        }
                        else if (uofDocType == Act.UofTranslator.UofTranslatorLib.DocType.Word)
                        {
                            implyDocType = DocumentType.Word;
                        }
                        else if (uofDocType == Act.UofTranslator.UofTranslatorLib.DocType.Powerpoint)
                        {
                            implyDocType = DocumentType.Powerpnt;
                        }
                        else
                        {
                            implyDocType = DocumentType.Unknown;
                        }
                    }
                    catch (Exception)
                    {
                        implyDocType = DocumentType.Unknown;
                    }
                }
                else
                {
                    implyDocType = value;
                }
            }
        }

        /// <summary>
        ///     Source File Name
        /// </summary>
        private string srcFileName;

        /// <summary>
        ///     Destination File Name
        /// </summary>
        private string dstFileName;

        /// <summary>
        ///     Present Progress
        ///     Can be either "Ready", "Working", "Done", "Paused"
        /// </summary>
        private int stateValProgress;

        public int StateValProgress
        {
            get { return stateValProgress; }
            set { stateValProgress = value; }
        }

        /// <summary>
        ///     Present Result
        ///     Can be "or" of "OK", "Error", "Warning"
        /// </summary>
        private int stateMaskResult;

        /// <summary>
        ///     Has Warning?
        ///     Check the Warning Mask.
        /// </summary>
        public bool HasWarning
        {
            get { return hasMask(StateVal.STATE_MASK_WARNING); }
            set { toggleMask(StateVal.STATE_MASK_WARNING, value); }
        }

        /// <summary>
        ///     Source File Name
        /// </summary>
        public string SourceFileName
        {
            get { return srcFileName; }
            set 
            {
                srcFileName = value;

                //We set the document Type by the SourceFileName
                if (File.Exists(srcFileName))
                {
                    string ext = Path.GetExtension(srcFileName);
                    if(ext.Equals(Extensions.OOX_WORD, StringComparison.OrdinalIgnoreCase))
                    {
                        docType = DocumentType.Word;
                        transType = TranslationType.OoxToUof;
                    }
                    else if(ext.Equals(Extensions.OOX_EXCEL, StringComparison.OrdinalIgnoreCase))
                    {
                        docType = DocumentType.Excel;
                        transType = TranslationType.OoxToUof;
                    }
                    else if(ext.Equals(Extensions.OOX_POWERPNT, StringComparison.OrdinalIgnoreCase))
                    {
                        docType = DocumentType.Powerpnt;
                        transType = TranslationType.OoxToUof;
                    }

                        // add by linyh
                    else if (ext.Equals(Extensions.UOF_EXCEL, StringComparison.OrdinalIgnoreCase))
                    {
                        docType = DocumentType.Excel;
                        transType = TranslationType.UofToOox;
                    }
                    else if (ext.Equals(Extensions.UOF_POWERPNT, StringComparison.OrdinalIgnoreCase))
                    {
                        docType = DocumentType.Powerpnt;
                        transType = TranslationType.UofToOox;
                    }
                    else if (ext.Equals(Extensions.UOF_WORD, StringComparison.OrdinalIgnoreCase))
                    {
                        docType = DocumentType.Word;
                        transType = TranslationType.UofToOox;
                    }
                    else
                    {
                        transType = TranslationType.Unknown;
                        docType = DocumentType.Unknown;
                    }
                }
            }
        }

        /// <summary>
        ///     Destination File Name
        /// </summary>
        public string DestinationFileName
        {
            get { return dstFileName; }
            set { dstFileName = value; }
        }

        /// <summary>
        ///     Is Translation Succeed?
        ///     Check the Error Mask
        /// </summary>
        public bool IsSucceed
        {
            get { return !(hasMask(StateVal.STATE_MASK_ERROR)); }
            set { toggleMask(StateVal.STATE_MASK_ERROR, (!value)); }
        }

        /// <summary>
        ///     Constructor of FileState Class
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        public FileState()
        {
            stateValProgress = StateVal.STATE_READY;
            docType = DocumentType.Unknown;
            transType = TranslationType.Unknown;
            implyDocType = DocumentType.Unknown;
            implyTransType = TranslationType.Unknown;
            srcFileName = null;
            dstFileName = null;
            IsSucceed = false;
            HasWarning = false;
        }

        public bool hasMask(int val)
        {
            return ((this.stateMaskResult & val) != 0);
        }

        public void toggleMask(int val, bool isSet)
        {
            if (isSet)
            {
                this.stateMaskResult = (this.stateMaskResult | val);
            }
            else
            {
                this.stateMaskResult = (this.stateMaskResult & (~val));
            }
        }

        public FileState clone()
        {
            FileState result = new FileState();
            result.srcFileName = this.srcFileName;
            result.dstFileName = this.dstFileName;
            result.stateMaskResult = this.stateMaskResult;
            result.stateValProgress = this.stateValProgress;
            result.docType = this.docType;
            result.transType = this.transType;
            result.implyDocType = this.implyDocType;
            result.implyTransType = this.implyTransType;
            return result;
        }

        public void setDirTranslationType(DocumentType docType, TranslationType transType)
        {
            this.docType = docType;
            this.transType = transType;
        }

        public void adjustDestFileExtension()
        {
            string Ext = getDestFileExtension();

            if (Ext == null)
            {
                return;
            }

            this.DestinationFileName = Path.GetDirectoryName(dstFileName) + Path.DirectorySeparatorChar + Path.GetFileNameWithoutExtension(dstFileName) + Ext;
        }

        public string getDestFileExtension()
        {
            string Ext = null;

            if (this.TransType == TranslationType.OoxToUof)
            {
               // Ext = Extensions.UOF;
                if (this.DocType == DocumentType.Word)
                {
                    Ext = Extensions.UOF_WORD;
                }
                else if (this.DocType == DocumentType.Excel)
                {
                    Ext = Extensions.UOF_EXCEL;
                }
                else if (this.DocType == DocumentType.Powerpnt)
                {
                    Ext = Extensions.UOF_POWERPNT;
                }
            }
            else if (this.TransType == TranslationType.UofToOox)
            {
                if (this.DocType == DocumentType.Word)
                {
                    Ext = Extensions.OOX_WORD;
                }
                else if (this.DocType == DocumentType.Excel)
                {
                    Ext = Extensions.OOX_EXCEL;
                }
                else if (this.DocType == DocumentType.Powerpnt)
                {
                    Ext = Extensions.OOX_POWERPNT;
                }
            }

            return Ext;
        }
    }
}
