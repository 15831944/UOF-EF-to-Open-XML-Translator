/*
 * Copyright (c) 2006, Beihang University, China
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

namespace Act.UofTranslator.UofTranslatorStrictLib
{
    /// <summary>
    /// This class is used to hold contants.
    /// </summary>
    /// <author>linwei</author>
    /// <modify>linyaohu</modify>
    internal class TranslatorConstants
    {
        public const string RESOURCE_LOCATION = "resources";

        public const string UOFToOOX_WORD_LOCATION = "word.uof2oox";

        public const string UOFToOOX_XSL = "uof2oox.xsl";

        public const string UOFToOOX_POSTTREAT_STEP1_XSL = "posttreatmentStep1.xsl";

        public const string UOFToOOX_COMPUTE_SIZE_XSL = "uof2oox-compute-size.xsl";

        public const string UOFToOOX_PRETREAT_STEP1_XSL = "preprocessing1.xsl";

        public const string UOFToOOX_PRETREAT_STEP2_XSL = "preprocessing2.xsl";

        public const string UOFToOOX_PRETREAT_SETP3_XSL = "preprocessing3.xsl";

        public const string OOXToUOF_WORD_LOCATION = "word.oox2uof";

        public const string OOXToUOF_PRETREAT_STEP1_XSL = "pretreatmentStep1.xsl";

        public const string OOXToUOF_PRETREAT_STEP2_XSL = "pretreatmentStep2.xsl";

        public const string OOXToUOF_PRETREAT_STEP3_XSL = "pretreatmentStep3.xsl";

        public const string OOXToUOF_XSL = "oox2uof.xsl";

        public const string OOXToUOF_COMPUTE_SIZE_XSL = "oox2uof-compute-size.xsl";

        public const string SOURCE_XML = "source.xml";

        public const string XSLFORTEST = "fortest.xsl";

        public const string MESSAGE_EN_RESX = "message-en";

        public const string MESSAGE_ZH_CHS_RESX = "message-zh-CHS";
        //added by lin yaohu
        public const string UOFToOOX_POWERPOINT_LOCATION = "Powerpoint.uof2oox";

        public const string UOFToOOX_EXCEL_LOCATION = "Excel.uof2oox";

        public const string OOXToUOF_POWERPOINT_LOCATION = "Powerpoint.oox2uof";

        public const string OOXToUOF_EXCEL_LOCATION = "Excel.oox2uof";
        //added by luo wentian
        //public const string UOF11ToUOF10_LOCATION = "uof11-10.xsl";
        //public const string UOF10ToUOF11_LOCATION = "uof10-11.xsl";

       // public const string XMLNS_UOFWORDPROC = "http://schemas.uof.org/cn/2009/uof-wordproc";
        public const string XMLNS_UOFWORDPROC = "http://schemas.uof.org/cn/2009/wordproc";

        public const string XMLNS_UOFSPREADSHEET = "http://schemas.uof.org/cn/2009/spreadsheet";

        public const string XMLNS_UOFPRESENTATION = "http://schemas.uof.org/cn/2009/presentation";

        public const string XMLNS_METADATA = "http://schemas.uof.org/cn/2009/metadata";

        public const string XMLNS_UOFOBJECTS = "http://schemas.uof.org/cn/2009/objects";

        public const string XMLNS_UOFRULES = "http://schemas.uof.org/cn/2009/rules";

        public const string XMLNS_UOFSTYLES = "http://schemas.uof.org/cn/2009/styles";

        public const string XMLNS_UOFGRAPH = "http://schemas.uof.org/cn/2009/graph";

        public const string XMLNS_UOFGRAPHICS = "http://schemas.uof.org/cn/2009/graphics";

        public const string XMLNS_UOFEXTEND = "http://schemas.uof.org/cn/2009/extend";

        public const string XMLNS_A = "http://purl.oclc.org/ooxml/drawingml/main";

        public const string XMLNS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

        public const string XMLNS_A15 = "http://schemas.microsoft.com/office/drawing/2012/main";

        public const string XMLNS_APP = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties";


        public const string XMLNS_C = "http://purl.oclc.org/ooxml/drawingml/chart";


        public const string XMLNS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";


        public const string XMLNS_CUS = "http://purl.oclc.org/ooxml/officeDocument/customProperties";


        public const string XMLNS_DC = "http://purl.org/dc/elements/1.1/";


        public const string XMLNS_DCMITYPE = "http://purl.org/dc/dcmitype/";


        public const string XMLNS_DCTERMS = "http://purl.org/dc/terms/";

        public const string XMLNS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";


        public const string XMLNS_FN = "http://www.w3.org/2005/xpath-functions";


        public const string XMLNS_FO = "http://www.w3.org/1999/XSL/Format";


        public const string XMLNS_M = "http://purl.oclc.org/ooxml/officeDocument/math";


        public const string XMLNS_O = "urn:schemas-microsoft-com:office:office";


        public const string XMLNS_O2UPIC = "urn:o2upic:xmlns:post-processings:special";


        public const string XMLNS_ORI = "http://purl.oclc.org/ooxml/officeDocument/relationships/image";


        public const string XMLNS_P = "http://purl.oclc.org/ooxml/presentationml/main";


        public const string XMLNS_PC = "http://schemas.openxmlformats.org/package/2006/content-types";


        public const string XMLNS_PIC = "http://purl.oclc.org/ooxml/drawingml/picture";


        public const string XMLNS_PR = "http://schemas.openxmlformats.org/package/2006/relationships";


        public const string XMLNS_PZIP = "urn:u2o:xmlns:post-processings:special";


        public const string XMLNS_R = "http://purl.oclc.org/ooxml/officeDocument/relationships";


        public const string XMLNS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";


        public const string XMLNS_UOF = "http://schemas.uof.org/cn/2009/uof";


        public const string XMLNS_V = "urn:schemas-microsoft-com:vml";


        public const string XMLNS_VE = "http://schemas.openxmlformats.org/markup-compatibility/2006";


        public const string XMLNS_VT = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes";


        public const string XMLNS_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";


        public const string XMLNS_W10 = "urn:schemas-microsoft-com:office:word";


        public const string XMLNS_WNE = "http://schemas.microsoft.com/office/word/2006/wordml";


        public const string XMLNS_WP = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";


        public const string XMLNS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";


        public const string XMLNS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";


        public const string XMLNS_WS = "http://purl.oclc.org/ooxml/spreadsheetml/main";


        public const string XMLNS_XDR = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";


        public const string XMLNS_XDT = "http://www.w3.org/2005/xpath-datatypes";


        public const string XMLNS_XS = "http://www.w3.org/2001/XMLSchema";


        public const string XMLNS_XSL = "http://www.w3.org/1999/XSL/Transform";

        






        public static string XMLNS_STYLES { get; set; }
    }
}
