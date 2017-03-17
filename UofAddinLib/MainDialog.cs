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
using System.Xml;
using System.Xml.Schema;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using System.Reflection;
using System.IO;
using log4net;
using Act.UofTranslator.UofTranslatorLib;
using Act.UofTranslator.UofTranslatorStrictLib;
using Act.UofTranslator.TranslatorMgr;

namespace UofAddinLib
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      Ö÷½çÃæ´°¿ÚÀà
    ///     Author:         µË×·
    ///     Create Date:    2007-05-23
    /// </summary>
    public partial class MainDialog : Form
    {

        #region Private Members

        /// <summary>
        ///     For Validation Error Logging
        /// </summary>
        private static ILog logger = LogManager.GetLogger("AddinLogger");

        /// <summary>
        ///     Is Log Now Being Viewed
        /// </summary>
        private bool isViewingLog;

        /// <summary>
        ///     The Height of the Log Textbox
        /// </summary>
        private int LogTextBoxHeight;

        /// <summary>
        ///     List for store the state of those files to be processed
        /// </summary>
        private List<FileState> FileStates;

        /// <summary>
        ///     The Type of this Translator, can be DialogType.Open or DialogType.Save
        /// </summary>
        private DialogType TranslatorType;

        /// <summary>
        ///     The Thread to perform translation
        /// </summary>
        private Thread TranslatorThread;

        /// <summary>
        ///     The Word Application
        /// </summary>
        private string[] filesOpenByWord;


        /// <summary>
        ///     If All File succeed in translation, this is true
        ///     Else false
        /// </summary>
        private bool isAllSucceed;

        /// <summary>
        ///     If Any File succeed in translation, this is true
        ///     Else false
        /// </summary>
        private bool isOneSucceed;

        /// <summary>
        ///     If Detail Box is not empty, this is true
        ///     Else false
        /// </summary>
        private bool isDetail;


        /// <summary>
        ///     If Translate is defered because of Handle not created
        /// </summary>
        private bool deferTranslate;

        /// <summary>
        ///     The Config Manager,which load and save the config
        /// </summary>
        private ConfigManager cfgManager;

        /// <summary>
        ///     The Uof Decompress Adapter
        /// </summary>
       // private Act.UofTranslator.UofTranslatorLib.IUofCompressAdapter Uof2OoxAdapter;

       // private Act.UofTranslator.UofTranslatorStrictLib.IUofCompressAdapter Uof2OoxStrictAdapter=null;

        /// <summary>
        ///     The Uof Compress Adapter
        /// </summary>
        private Act.UofTranslator.UofTranslatorLib.IUofCompressAdapter Oox2UofAdapter;

        private Act.UofTranslator.UofTranslatorStrictLib.IUofCompressAdapter Oox2UofStrictAdapter=null;

        /// <summary>
        ///     Indicate whether the adapter config file is valid
        /// </summary>
        private bool isConfigValid;

        /// <summary>
        ///     Delegate for Set Progress Bar
        /// </summary>
        private delegate void SetProgressBarDelegate();

        /// <summary>
        ///     Delegate for Set Current State Label
        /// </summary>
        /// <param name="state">
        ///     Current State
        /// </param>
        private delegate void SetCurrentStateDelegate(string state);

        /// <summary>
        ///     Delegate for Write Log
        /// </summary>
        /// <param name="log">
        ///     Log String
        /// </param>
        private delegate void WriteLogDelegate(string log);

        /// <summary>
        ///     Delegate for Close Form
        /// </summary>
        private delegate void CloseFormDelegate();

        /// <summary>
        ///     Delegate for Disable Cancel Button
        /// </summary>
        private delegate void DisableCancelButtonDelegate();

        /// <summary>
        ///     Delegate for Setting cancel button to exit
        /// </summary>
        private delegate void SetCancelButtonExitDelegate();

        /// <summary>
        ///     Function Pointer of the SetProgressBarMoveDelegate
        /// </summary>
        private SetProgressBarDelegate setProgressBarMoveDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetProgressBarStopDelegate
        /// </summary>
        private SetProgressBarDelegate setProgressBarStopDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetProgressBarDelegate
        /// </summary>
        private SetProgressBarDelegate setProgressBarDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetProgressBarDelegate, But For Reset the Progress Bar
        /// </summary>
        private SetProgressBarDelegate resetProgressBarDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetCurrentStateDelegate
        /// </summary>
        private SetCurrentStateDelegate setCurrentStateDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetOverAllFileDelegate
        /// </summary>
        private SetCurrentStateDelegate setOverAllFileDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetOverAllFileDelegate
        /// </summary>
        private SetCurrentStateDelegate setToFileDelegateCall;

        /// <summary>
        ///     Function Pointer of the WriteLogDelegate
        /// </summary>
        private WriteLogDelegate writeLogDelegateCall;

        /// <summary>
        ///     Function Pointer of the CloseFormDelegate
        /// </summary>
        private CloseFormDelegate closeFormDelegateCall;

        /// <summary>
        ///     Function Pointer of the DisableCancelButtonDelegate
        /// </summary>
        private DisableCancelButtonDelegate disableCancelButtonDelegateCall;

        /// <summary>
        ///     Function Pointer of the SetCancelButtonExitDelegate
        /// </summary>
        private SetCancelButtonExitDelegate setCancelButtonExitDelegateCall;

        /// <summary>
        ///     Delegate for the Cancel Button
        /// </summary>
        private EventHandler cancelButtonEventHandler;

        /// <summary>
        ///     Delegate for Process the Progress Changed Event
        /// </summary>
        private EventHandler progressListener;

        /// <summary>
        ///     Delegate for Process the Feedback Event, For Displaying at Status Label
        /// </summary>
        private EventHandler feedbackListenerForLabel;

        /// <summary>
        ///     Delegate for Process the Feedback Event, For Displaying at Log Text box
        /// </summary>
        private EventHandler feedbackListenerForLog;

        #endregion


        /// <summary>
        ///     Current File's States
        /// </summary>
        public List<FileState> CurrentFileStates
        {
            get
            {
                return FileStates;
            }
        }

        /// <summary>
        ///     Constructor of MainDialog Class
        /// </summary>
        /// <param name="wordOpeningFiles">
        ///     The File currently being open by word.
        ///     Only Used when called in word add-in, else set it to null;
        /// </param>
        /// <param name="type">
        ///     The Dialog Type, Can be DialogType.Open or DialogType.Save
        /// </param>
        /// <param name="docType">
        ///     The Document Type, Which can be "Word", "Excel" or "Powerpnt"
        /// </param>
        /// <param name="selectedFiles">
        ///     The Selected Files To be Translated
        /// </param>
        /// <param name="tmpSavedDocxFile">
        ///     The Full Path of the temp saved docx file
        ///     Only used if type is DialogType.Save
        /// </param>
        /// <param name="targetFiles">
        ///     The Target Files of Translation
        ///     Only Used When the type is DialogType.UofToOox or DialogType.OoxToUof
        /// </param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        public MainDialog(string[] wordOpeningFiles, DialogType type, DocumentType docType, string[] selectedFiles, string tmpSavedDocxFile, string[] targetFiles)
        {
            FileState currentFileState = null;
            string fileName = null;

            Application.EnableVisualStyles();

            // Set Type, applicationObject
            TranslatorType = type;
            filesOpenByWord = wordOpeningFiles;
            isAllSucceed = false;
            isOneSucceed = false;
            isDetail = false;

            // Initialize the Config Manager
            try
            {
                cfgManager = new ConfigManager(System.IO.Path.GetDirectoryName(typeof(MainDialog).Assembly.Location) + @"\conf\config.xml");
                cfgManager.LoadConfig();
            }
            catch (Exception)
            {
                // Do nothing.
            }

            // Find the Adapter
            FindAdapter();

            InitializeComponent();

            LogTextBoxHeight = txtLog.Height;

            // Set the control based on the config
            chkbxIgnoreError.Checked = cfgManager.IsErrorIgnored;

            // Set Log TextBox Unvisible
            isViewingLog = false;
            HideLog();

            // Get All Selected Files' name
            FileStates = new List<FileState>();
            for (int i = 0; i < selectedFiles.Length; i++)
            {
                fileName = selectedFiles[i];
                currentFileState = new FileState();
                if (type == DialogType.Open)
                {
                    // If it's open, The File Name is Source
                    currentFileState.SourceFileName = fileName;
                    currentFileState.ImplyTransType = TranslationType.UofToOox;
                    currentFileState.ImplyDocType = docType;
                }
                else if (type == DialogType.Save)
                {
                    // If it's save, The File Name is Target
                    currentFileState.SourceFileName = tmpSavedDocxFile;
                    currentFileState.DestinationFileName = fileName;
                    currentFileState.ImplyTransType = TranslationType.OoxToUof;
                    currentFileState.ImplyDocType = docType;
                }
                else if ((type == DialogType.OoxToUof) || (type == DialogType.UofToOox))
                {
                    // If it's conversion, The source and target is done.
                    currentFileState.SourceFileName = fileName;
                    currentFileState.DestinationFileName = targetFiles[i];

                    if (type == DialogType.OoxToUof)
                    {
                        currentFileState.ImplyTransType = TranslationType.OoxToUof;
                    }
                    else
                    {
                        currentFileState.ImplyTransType = TranslationType.UofToOox;
                    }
                    currentFileState.ImplyDocType = docType;
                }
                FileStates.Add(currentFileState);
            }

            // Nullize other variables
            TranslatorThread = null;

            // Set Delegates
            setProgressBarMoveDelegateCall = new SetProgressBarDelegate(SetProgressBarMoveCore);
            setProgressBarStopDelegateCall = new SetProgressBarDelegate(SetProgressBarStopCore);
            setProgressBarDelegateCall = new SetProgressBarDelegate(SetProgressBarCore);
            resetProgressBarDelegateCall = new SetProgressBarDelegate(ResetProgressBarCore);
            setCurrentStateDelegateCall = new SetCurrentStateDelegate(SetCurrentStateCore);
            setOverAllFileDelegateCall = new SetCurrentStateDelegate(SetOverAllFileCore);
            setToFileDelegateCall = new SetCurrentStateDelegate(SetToFileCore);
            writeLogDelegateCall = new WriteLogDelegate(WriteLogCore);
            closeFormDelegateCall = new CloseFormDelegate(CloseFormCore);
            disableCancelButtonDelegateCall = new DisableCancelButtonDelegate(DisableCancelButtonCore);
            progressListener = new EventHandler(SetProgressBar);
            feedbackListenerForLabel = new EventHandler(SetCurrentState);
            feedbackListenerForLog = new EventHandler(WriteLog);
            setCancelButtonExitDelegateCall = new SetCancelButtonExitDelegate(SetCancelButtonExitCore);

            cancelButtonEventHandler = new EventHandler(btnCancel_Click);
            this.btnCancel.Click += cancelButtonEventHandler;


            // Start Translate
            StartTranslate();
        }

        private void UofToOoxComputeSize(string fileName)
        {
            //TODO:Èç¹ûComputeSizeºÃÁËµÄ»°ÔÚÕâÀï²¹ÉÏ¡£¡£¡£
        }

        private void OoxToUofComputeSize(string fileName)
        {
            //TODO:Èç¹ûComputeSizeºÃÁËµÄ»°ÔÚÕâÀï²¹ÉÏ¡£¡£¡£
        }

        /// <summary>
        ///     Hide The Log TextBox and Resize the Dialog
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        private void HideLog()
        {

            // Resize the Dialog
            this.Height -= LogTextBoxHeight;

            // Set the Log Text invisible
            txtLog.Visible = false;
        }

        /// <summary>
        ///     Show The Log TextBox and Resize the Dialog
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        private void ShowLog()
        {
            // Resize the Dialog
            this.Height += LogTextBoxHeight;

            // Set the Log Text invisible
            txtLog.Visible = true;
        }

        /// <summary>
        ///     Start the Translation Thread
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void StartTranslate()
        {
            // Start the search if the handle has
            // been created. Otherwise, defer it until the
            // handle has been created.
            if (IsHandleCreated)
            {
                TranslatorThread = new Thread(new ThreadStart(TranslateRoutine));
                TranslatorThread.Start();
            }
            else
            {
                deferTranslate = true;
            }

        }

        /// <summary>
        ///     Terminate the Translation Thread
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void Cancel()
        {
            if (TranslatorThread != null)
            {
                if (TranslatorThread.IsAlive)
                {
                    TranslatorThread.Abort();
                    TranslatorThread.Join();
                }

                TranslatorThread = null;
            }
        }

        /// <summary>
        /// initail transitional converter lib
        /// </summary>
        /// <param name="inputFileName">input file</param>
        /// <param name="type">document type</param>
        /// <returns>transitional converter interface</returns>
        private Act.UofTranslator.UofTranslatorLib.IUOFTranslator initConverter(string inputFileName, DocumentType type)
        {
            Act.UofTranslator.UofTranslatorLib.IUOFTranslator Converter;
            if (type == DocumentType.Word)
            {
                Converter = Act.UofTranslator.UofTranslatorLib.TranslatorFactory.CheckFileType(Act.UofTranslator.UofTranslatorLib.DocType.Word);
                Converter.AddProgressMessageListener(this.progressListener);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLabel);
                return Converter;
            }
            else if (type == DocumentType.Excel)
            {
                Converter = Act.UofTranslator.UofTranslatorLib.TranslatorFactory.CheckFileType(Act.UofTranslator.UofTranslatorLib.DocType.Excel);
                Converter.AddProgressMessageListener(this.progressListener);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLabel);
                return Converter;
            }
            else if (type == DocumentType.Powerpnt)
            {
                Converter = Act.UofTranslator.UofTranslatorLib.TranslatorFactory.CheckFileType(Act.UofTranslator.UofTranslatorLib.DocType.Powerpoint);
                Converter.AddProgressMessageListener(this.progressListener);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLabel);
                return Converter;
            }

            return null;
        }

        /// <summary>
        /// initial strict converter lib
        /// </summary>
        /// <param name="inputFileName">input file</param>
        /// <param name="type">document type</param>
        /// <returns>strict converter interface</returns>
        private Act.UofTranslator.UofTranslatorStrictLib.IUOFTranslator initStrictConverter(string inputFileName, DocumentType type)
        {
            Act.UofTranslator.UofTranslatorStrictLib.IUOFTranslator strictConverter;
            if (type == DocumentType.Word)
            {
                strictConverter = Act.UofTranslator.UofTranslatorStrictLib.TranslatorFactory.CheckFileType(Act.UofTranslator.UofTranslatorStrictLib.DocType.Word);
                strictConverter.AddProgressMessageListener(this.progressListener);
                strictConverter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                strictConverter.AddFeedbackMessageListener(this.feedbackListenerForLabel);
                return strictConverter;
            }
            else if (type == DocumentType.Excel)
            {
                strictConverter = Act.UofTranslator.UofTranslatorStrictLib.TranslatorFactory.CheckFileType(Act.UofTranslator.UofTranslatorStrictLib.DocType.Excel);
                strictConverter.AddProgressMessageListener(this.progressListener);
                strictConverter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                strictConverter.AddFeedbackMessageListener(this.feedbackListenerForLabel);
                return strictConverter;
            }
            else if (type == DocumentType.Powerpnt)
            {
                strictConverter = Act.UofTranslator.UofTranslatorStrictLib.TranslatorFactory.CheckFileType(Act.UofTranslator.UofTranslatorStrictLib.DocType.Powerpoint);
                strictConverter.AddProgressMessageListener(this.progressListener);
                strictConverter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                strictConverter.AddFeedbackMessageListener(this.feedbackListenerForLabel);
                return strictConverter;
            }

            return null;
        }

        /// <summary>
        ///     The Main Translation Routine
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void TranslateRoutine()
        {
            Act.UofTranslator.UofTranslatorLib.IUOFTranslator converter = null;
            Act.UofTranslator.UofTranslatorStrictLib.IUOFTranslator strictConverter = null;

            DialogType type;
            string outputFileName;
            int currentFileIndex;
            int SucceedCount;
            bool overwriteAll=false;
            bool skipAll=false;
            try
            {
                //Test if File Array is null
                if (FileStates.Count == 0)
                {
                    BeginInvoke(closeFormDelegateCall);
                    return;
                }

                type = TranslatorType;
                outputFileName = null;
                currentFileIndex = 0;
                isAllSucceed = false;
                isOneSucceed = false;
                overwriteAll = false;
                skipAll = false;

                //Set the marquee progress bar move
                BeginInvoke(setProgressBarMoveDelegateCall);

                foreach (FileState currentFileState in FileStates)
                {
                    // Clean The Progress Bar
                    Invoke(resetProgressBarDelegateCall);

                    currentFileIndex++;

                    // Display Current Processing File
                    BeginInvoke(setOverAllFileDelegateCall, new object[] {"(" + currentFileIndex.ToString() + 
                "/" + FileStates.Count + ")" + System.IO.Path.GetFileName(currentFileState.SourceFileName)});

                    #region Open UOF

                    if (type == DialogType.Open)
                    {
                        // Make the New File Name
                        outputFileName = System.IO.Path.GetTempPath()
                                        + System.IO.Path.GetFileNameWithoutExtension(currentFileState.SourceFileName);
                        outputFileName = MakeFileName(outputFileName, currentFileState.getDestFileExtension());

                        if (outputFileName != null)
                        {
                            // All Done, Start the Translation
                            BeginInvoke(setToFileDelegateCall, new object[] { System.IO.Path.GetFileName(outputFileName) });

                            currentFileState.DestinationFileName = outputFileName;
                            try
                            {
                                {
                                    if (cfgManager.IsUofToTransitioanlOOX) // to transitional
                                    {
                                        converter = this.initConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                        if (converter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        UofToOoxComputeSize(currentFileState.SourceFileName);
                                        converter.UofToOox(currentFileState.SourceFileName, outputFileName);
                                    }
                                    else // to strict
                                    {
                                        strictConverter=this.initStrictConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                        if (strictConverter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        //UofToOoxComputeSize(currentFileState.SourceFileName);
                                        strictConverter.UofToOox(currentFileState.SourceFileName, outputFileName);

                                    }
                                }
                                currentFileState.IsSucceed = true;
                            }
                            catch (Exception e)
                            {
                                BeginInvoke(writeLogDelegateCall, new object[] { "Error:" + e.Message });
                                currentFileState.IsSucceed = false;
                            }
                            
                        }
                        else
                        {
                            // File Name Error, Go next
                            BeginInvoke(setToFileDelegateCall, new object[] { "" });

                            currentFileState.IsSucceed = false;
                            continue;
                        }
                    }

                    #endregion

                    #region Save As UOF

                    else if (type == DialogType.Save)
                    {
                        if (currentFileState.DestinationFileName != null)
                        {
                            // All Done, Start the Translation
                            BeginInvoke(setToFileDelegateCall, new object[] { System.IO.Path.GetFileName(currentFileState.DestinationFileName) });

                            try
                            {
                                 Act.UofTranslator.TranslatorMgr.MSDocType msDocType = Act.UofTranslator.TranslatorMgr.TranslatorMgr.CheckMSFileType(currentFileState.SourceFileName);
                                    if (msDocType == Act.UofTranslator.TranslatorMgr.MSDocType.StrictExcel || msDocType == Act.UofTranslator.TranslatorMgr.MSDocType.StrictPowerpoint ||
                                         msDocType == Act.UofTranslator.TranslatorMgr.MSDocType.StrictWord)
                                    {
                                        strictConverter = this.initStrictConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                      
                                        if (strictConverter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        OoxToUofComputeSize(currentFileState.SourceFileName);
                                        strictConverter.OoxToUof(currentFileState.SourceFileName, currentFileState.DestinationFileName);

                                    }
                                    else
                                    {
                                        converter = this.initConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                        if (converter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        OoxToUofComputeSize(currentFileState.SourceFileName);
                                        converter.OoxToUof(currentFileState.SourceFileName, currentFileState.DestinationFileName);
                                    }
                                    currentFileState.IsSucceed = true;
                               
                            }
                            catch (Exception e)
                            {
                                BeginInvoke(writeLogDelegateCall, new object[] { "Error:" + e.Message });
                                currentFileState.IsSucceed = false;
                            }
                            finally
                            {
                                try
                                {
                                    // Clear Adapter Temp File
                                    if (Oox2UofAdapter != null)
                                    {
                                        if (File.Exists(Oox2UofAdapter.OutputFilename) && (!Oox2UofAdapter.OutputFilename.Equals(currentFileState.DestinationFileName)))
                                        {
                                            File.Delete(Oox2UofAdapter.OutputFilename);
                                        }
                                    }
                                    if (Oox2UofStrictAdapter != null)
                                    {
                                        if (File.Exists(Oox2UofStrictAdapter.OutputFilename) && (!Oox2UofStrictAdapter.OutputFilename.Equals(currentFileState.DestinationFileName)))
                                        {
                                            File.Delete(Oox2UofStrictAdapter.OutputFilename);
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    logger.Warn("Fail To Delete Adapter's Temp File: " + Oox2UofAdapter.OutputFilename+e.Message);
                                }
                            }
                        }
                        else
                        {
                            // File Name Error, Go next
                            BeginInvoke(setToFileDelegateCall, new object[] { "" });

                            currentFileState.IsSucceed = false;
                            continue;
                        }
                    }

                    #endregion

                    else if ((type == DialogType.UofToOox) || (type == DialogType.OoxToUof))
                    {
                       // bool donecurrent = false;
                       // bool skipped = false;

                         // System.Windows.Forms.MessageBox.Show("Destination File=" + currentFileState.DestinationFileName);

                        // If Only One File, Skip Check
                        if (FileStates.Count > 1)
                        {


                            if (File.Exists(currentFileState.DestinationFileName))
                            {
                                string currentFileName = currentFileState.DestinationFileName;
                                currentFileState.DestinationFileName = MakeFileName(Path.GetDirectoryName(currentFileName) + "\\" + Path.GetFileNameWithoutExtension(currentFileName), currentFileState.getDestFileExtension());
                               // System.Windows.Forms.MessageBox.Show(currentFileState.DestinationFileName);
                            }
                            
                        }

                        // Start the Translation
                        BeginInvoke(setToFileDelegateCall, new object[] { System.IO.Path.GetFileName(currentFileState.DestinationFileName) });

                        #region UOF To OOX

                        if (type == DialogType.UofToOox)
                        {
                            try
                            {                                
                                    // to transitional
                                    if (cfgManager.IsUofToTransitioanlOOX)
                                    {
                                        converter = this.initConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                        if (converter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        UofToOoxComputeSize(currentFileState.SourceFileName);
                                        converter.UofToOox(currentFileState.SourceFileName, currentFileState.DestinationFileName);
                                    }
                                    else // to strict
                                    {
                                        strictConverter = this.initStrictConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                        if (strictConverter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                      //  UofToOoxComputeSize(Uof2OoxAdapter.OutputFilename);
                                        strictConverter.UofToOox(currentFileState.SourceFileName, currentFileState.DestinationFileName);
                                    }
                                
                                currentFileState.IsSucceed = true;
                            }
                            catch (Exception e)
                            {
                                BeginInvoke(writeLogDelegateCall, new object[] { "Error:" + e.Message });
                                currentFileState.IsSucceed = false;
                            }
                           
                        }

                        #endregion

                        #region OOX TO UOF

                        else
                        {
                             try
                            {
                               // System.Windows.Forms.MessageBox.Show("oox2uofAdapt=" + Oox2UofAdapter + "oox2uofstrictadapt=" + Oox2UofStrictAdapter);
                                  
                                //if (Oox2UofAdapter != null || Oox2UofStrictAdapter!=null)
                                //{
                                    
                                    // select strict converter or transitional converter
                                    Act.UofTranslator.TranslatorMgr.MSDocType msDocType = Act.UofTranslator.TranslatorMgr.TranslatorMgr.CheckMSFileType(currentFileState.SourceFileName);
                                    if (msDocType == Act.UofTranslator.TranslatorMgr.MSDocType.StrictExcel || msDocType == Act.UofTranslator.TranslatorMgr.MSDocType.StrictPowerpoint ||
                                         msDocType == Act.UofTranslator.TranslatorMgr.MSDocType.StrictWord)
                                    {
                                        strictConverter = this.initStrictConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                       // System.Windows.Forms.MessageBox.Show("strictConverter=" + strictConverter);
                                   
                                        if (strictConverter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        OoxToUofComputeSize(currentFileState.SourceFileName);
                                        strictConverter.OoxToUof(currentFileState.SourceFileName, currentFileState.DestinationFileName);

                                        //Oox2UofStrictAdapter.InputFilename = currentFileState.DestinationFileName;

                                        //if (Oox2UofStrictAdapter.Transform())
                                        //{
                                        //    string tempfile = Path.GetTempFileName();
                                        //    File.Copy(Oox2UofStrictAdapter.OutputFilename, tempfile, true);
                                        //    File.Copy(tempfile, currentFileState.DestinationFileName, true);
                                        //    File.Delete(tempfile);
                                        //    Oox2UofStrictAdapter.OutputFilename = null;
                                        //}
                                        //else
                                        //{
                                        //    // Tranform failed
                                        //}
                                    }
                                    
                                //}
                                    else
                                    {
                                        converter = this.initConverter(currentFileState.SourceFileName, currentFileState.DocType);

                                        if (converter == null)
                                        {
                                            currentFileState.IsSucceed = false;
                                            continue;
                                        }
                                        OoxToUofComputeSize(currentFileState.SourceFileName);
                                        converter.OoxToUof(currentFileState.SourceFileName, currentFileState.DestinationFileName);
                                    }
                                    currentFileState.IsSucceed = true;
                            }
                            catch (Exception e)
                            {
                                BeginInvoke(writeLogDelegateCall, new object[] { "Error:" + e.Message });
                                currentFileState.IsSucceed = false;
                            }
                            finally
                            {
                                try
                                {
                                    // Clear Adapter Temp File
                                    if (Oox2UofAdapter != null)
                                    {
                                        if (File.Exists(Oox2UofAdapter.OutputFilename) && (!Oox2UofAdapter.OutputFilename.Equals(currentFileState.DestinationFileName)))
                                        {
                                            File.Delete(Oox2UofAdapter.OutputFilename);
                                        }
                                    }
                                    if (Oox2UofStrictAdapter != null)
                                    {
                                        if (File.Exists(Oox2UofStrictAdapter.OutputFilename) && (!Oox2UofStrictAdapter.OutputFilename.Equals(currentFileState.DestinationFileName)))
                                        {
                                            File.Delete(Oox2UofStrictAdapter.OutputFilename);
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    logger.Warn("Fail To Delete Adapter's Temp File: " + Oox2UofAdapter.OutputFilename+e.Message);
                                }
                            }
                        }

                        #endregion
                    }
                }

                // Disable Cancel Button
                BeginInvoke(disableCancelButtonDelegateCall);

                //Set the marquee progress bar Stop
                BeginInvoke(setProgressBarStopDelegateCall);

                // Now Translate Over
                if ((type == DialogType.Open) || (type == DialogType.UofToOox) || (type == DialogType.OoxToUof))
                {
                    SucceedCount = 0;

                    // Check if All Files Are Done.
                    foreach (FileState currentFileState in FileStates)
                    {
                        if (currentFileState.IsSucceed)
                        {
                            SucceedCount++;
                            isOneSucceed = true;
                        }
                    }

                    if (((SucceedCount < FileStates.Count) || (isDetail)) && (!cfgManager.IsErrorIgnored))
                    {
                        if (SucceedCount < FileStates.Count)
                        {
                            // Not All Files Are Translated
                            string msgError = String.Format(TranslatorRes.ResourceManager.GetString("OpenFileWarning"),
                                                        SucceedCount, FileStates.Count);
                            MessageBox.Show(msgError, TranslatorRes.ResourceManager.GetString("OpenFileWarningTitle"),
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            isAllSucceed = false;
                            BeginInvoke(setCurrentStateDelegateCall, TranslatorRes.ResourceManager.GetString("EndPrompt"));
                        }
                        if (isDetail)
                        {
                            BeginInvoke(setCurrentStateDelegateCall, TranslatorRes.ResourceManager.GetString("ForDetailPrompt"));
                        }
                    }
                    else
                    {
                        if (SucceedCount < FileStates.Count)
                        {
                            isAllSucceed = false;
                        }
                        else
                        {
                            isAllSucceed = true;
                        }
                        Invoke(closeFormDelegateCall);
                    }


                    // Now Open the translated Files in Word
                }
                else if (type == DialogType.Save)
                {
                    SucceedCount = 0;

                    // Check if All Files Are Done.
                    foreach (FileState currentFileState in FileStates)
                    {
                        if (currentFileState.IsSucceed)
                        {
                            SucceedCount++;
                            isOneSucceed = true;
                        }
                    }

                    if (((SucceedCount < FileStates.Count) || (isDetail)) && (!cfgManager.IsErrorIgnored))
                    {
                        if (SucceedCount < FileStates.Count)
                        {
                            // Not All Files Are Translated
                            string msgError = String.Format(TranslatorRes.ResourceManager.GetString("SaveFileError"),
                                                        SucceedCount, FileStates.Count);
                            MessageBox.Show(msgError, TranslatorRes.ResourceManager.GetString("SaveFileErrorTitle"),
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                            isAllSucceed = false;
                            BeginInvoke(setCurrentStateDelegateCall, TranslatorRes.ResourceManager.GetString("EndPrompt"));
                        }
                        if (isDetail)
                        {
                            BeginInvoke(setCurrentStateDelegateCall, TranslatorRes.ResourceManager.GetString("ForDetailPrompt"));
                        }
                    }
                    else
                    {
                        if (SucceedCount < FileStates.Count)
                        {
                            isAllSucceed = false;
                        }
                        else
                        {
                            isAllSucceed = true;
                        }
                        Invoke(closeFormDelegateCall);
                    }
                }

                // If now the window is not closed, change the cancel button's text and function
                BeginInvoke(setCancelButtonExitDelegateCall);
            }
            catch (Exception e)
            {
                if (e is ThreadAbortException)
                {
                    // Aborted, Do Nothing
                    BeginInvoke(setProgressBarStopDelegateCall);
                }
                else
                {
                    BeginInvoke(setProgressBarStopDelegateCall);
                    System.Windows.Forms.MessageBox.Show(e.ToString());
                }
            }
        }


        /// <summary>
        ///     Make A Docx File Name Based on the Original File Name
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
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private string MakeFileName(string baseFileName, string extendFileName)
        {
            string outputFileName = baseFileName;
            string extFileName = extendFileName;
            List<string> currentOpening = new List<string>();
            UInt64 count = 1;
            bool isOK = false;

            if (filesOpenByWord != null)
            {
                // Get All File Opening First
                foreach (string doc in filesOpenByWord)
                {
                    currentOpening.Add(doc);
                }
            }

            // And Add All Succecced File's Destination too
            foreach (FileState currentFileState in FileStates)
            {
                if (currentFileState.IsSucceed)
                {
                    currentOpening.Add(currentFileState.DestinationFileName);
                }
            }

           // System.Windows.Forms.MessageBox.Show("outputfileName=" + outputFileName);

            while (!isOK)
            {
                isOK = true;

                if (File.Exists(outputFileName + "(" + count.ToString() + ")" + extFileName))
                {
                    isOK = false;
                    count++;
                    if (count == System.UInt64.MaxValue)
                    {
                        // To Many Files, We Have an error
                        // This File Can't be Create, return null
                        return null;
                    }
                }
                //foreach (string fileName in currentOpening)
                //{
                // The Temp File Already Exist, Think about another name


                //   System.Windows.Forms.MessageBox.Show("compare=" + outputFileName + "(" + count.ToString() + ")" + extFileName);
                //    if (outputFileName.Equals(, StringComparison.OrdinalIgnoreCase))
                //    {

                //        break;
                //    }
                ////}
                if (isOK)
                {
                    try
                    {
                        if (System.IO.File.Exists(outputFileName + "(" + count.ToString() + ")" + extFileName))
                        {
                            // If the file already exist,test if it can be opened exclusively
                            System.IO.FileStream fs = System.IO.File.Open(outputFileName + "(" + count.ToString() + ")" + extFileName,
                                                                        System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite,
                                                                        System.IO.FileShare.None);
                            fs.Close();

                           // System.Windows.Forms.MessageBox.Show("outputfileName2=" + outputFileName);
                        }
                    }
                    catch (Exception)
                    {
                        // Oh We'll Change a name
                        count++;
                        isOK = false;
                    }
                }
            }

            outputFileName = outputFileName + "(" + count.ToString() + ")" + extFileName;
            return outputFileName;
        }

        /// <summary>
        /// for Bat processing files
        /// </summary>
        /// <param name="baseFileName">base file</param>
        /// <param name="extendFileName">extend File</param>
        /// <returns></returns>
        private string MakeBatFileName(string baseFileName, string extendFileName)
        {
            string newFileName = string.Empty;
            // like XXX(X).docx
            if (baseFileName.Contains(")."))
            {
                string count = baseFileName.Substring(baseFileName.IndexOf("("), baseFileName.IndexOf(")") - baseFileName.IndexOf("(") - 1);
                System.Windows.Forms.MessageBox.Show("count=" + count);
                int currentCount = 1;
                try
                {
                    currentCount = Convert.ToInt32(count) + 1;
                    newFileName = baseFileName.Substring(0, baseFileName.IndexOf("(")) + "(" + currentCount + ")" + extendFileName;
                }
                catch (Exception ex)
                {
                    newFileName = baseFileName + "(1)" + extendFileName;
                    logger.Error(ex.Message);
                }

            }
            else
            {
                newFileName = baseFileName + "(1)" + extendFileName;
            }

            return newFileName;
        }

        /// <summary>
        ///     Read Config File And Build Adapter
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-11-21
        private void FindAdapter()
        {
            XmlReader configReader = null;
            XmlSchemaSet sc = null;
            Stream xsdStream = null;
            Assembly asm = null;
            string uof2ooxAssemblyPath = null;
            string oox2uofAssemblyPath = null;
            string uof2ooxAdapterClassName = null;
            string oox2uofAdapterClassName = null;

            isConfigValid = true;

            try
            {
                // Get Current Assembly and XSD Stream
                asm = Assembly.GetExecutingAssembly();

                foreach (string name in asm.GetManifestResourceNames())
                {
                    // Find Resource in the Assembly
                    if (name.EndsWith("UofCompressAdapter.xsd"))
                    {
                        xsdStream = asm.GetManifestResourceStream(name);
                        break;
                    }
                }

                if (xsdStream == null)
                {
                    return;
                }

                // Build the Schema Set
                sc = new XmlSchemaSet();
                sc.Add(null, new XmlTextReader(xsdStream));

                // Try to Read the Config File: adatper.config
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.ValidationType = ValidationType.Schema;
                settings.Schemas = sc;
                settings.ValidationEventHandler += new ValidationEventHandler(AdapterConfigValidationCallBack);

                configReader = XmlReader.Create(Path.GetDirectoryName(asm.Location) + Path.DirectorySeparatorChar + "adapter.config", settings);

                while (configReader.Read())
                {
                    switch (configReader.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (configReader.Name.Equals("Uof2OoxAdapter"))
                            {
                                if (!configReader.ReadToDescendant("AssemblyName"))
                                {
                                    continue;
                                }
                                uof2ooxAssemblyPath = configReader.ReadString();
                                if (!configReader.ReadToNextSibling("ClassName"))
                                {
                                    continue;
                                }
                                uof2ooxAdapterClassName = configReader.ReadString();
                            }
                            else if (configReader.Name.Equals("Oox2UofAdapter"))
                            {
                                if (!configReader.ReadToDescendant("AssemblyName"))
                                {
                                    continue;
                                }
                                oox2uofAssemblyPath = configReader.ReadString();
                                if (!configReader.ReadToNextSibling("ClassName"))
                                {
                                    continue;
                                }
                                oox2uofAdapterClassName = configReader.ReadString();
                            }
                            break;
                        default:
                            break;
                    }
                }

                if (!isConfigValid)
                {
                    // Halt here
                    throw new Exception("Config File Not Valid! Stop Loading Adapter...");
                }

              //  Uof2OoxAdapter = (Act.UofTranslator.UofTranslatorLib.IUofCompressAdapter)(Assembly.LoadFrom(uof2ooxAssemblyPath).CreateInstance(uof2ooxAdapterClassName));
                Oox2UofAdapter = (Act.UofTranslator.UofTranslatorLib.IUofCompressAdapter)(Assembly.LoadFrom(oox2uofAssemblyPath).CreateInstance(oox2uofAdapterClassName));

               // Uof2OoxStrictAdapter = (Act.UofTranslator.UofTranslatorStrictLib.IUofCompressAdapter)(Assembly.LoadFrom(uof2ooxAssemblyPath).CreateInstance(uof2ooxAdapterClassName));
               // Oox2UofStrictAdapter = (Act.UofTranslator.UofTranslatorStrictLib.IUofCompressAdapter)(Assembly.LoadFrom(oox2uofAssemblyPath).CreateInstance(oox2uofAdapterClassName));



            }
            catch (Exception e)
            {
                logger.Error("Error While Loading Adapter: " + e.Message);
              //  Uof2OoxAdapter = null;
                Oox2UofAdapter = null;
            }
            finally
            {
                if (configReader != null)
                {
                    configReader.Close();
                }
            }
        }

        /// <summary>
        ///     Callback Function for Adapter Config File Validation
        /// </summary>
        /// <param name="sender">The Event Sender</param>
        /// <param name="e">The Event Parameter</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-11-21
        private void AdapterConfigValidationCallBack(object sender, ValidationEventArgs e)
        {
            isConfigValid = false;
            logger.Error(e.Message);
        }


        #region Delegate for Cross-thread call

        /// <summary>
        ///     This Function Makes the ProgressBar to Move if it is marquee style
        /// </summary>
        private void SetProgressBarMoveCore()
        {
            this.pbrCurrentFile.MarqueeAnimationSpeed = 100;
        }

        /// <summary>
        ///     This Function Makes the ProgressBar Stop to Move if it is marquee style
        /// </summary>
        private void SetProgressBarStopCore()
        {
            this.pbrCurrentFile.MarqueeAnimationSpeed = 0;
        }

        /// <summary>
        ///     This Function Increment the Progress Bar's Current Value By One
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void SetProgressBarCore()
        {
            if (pbrCurrentFile.Value < pbrCurrentFile.Maximum)
            {
                pbrCurrentFile.Value += 1;
            }
        }

        /// <summary>
        ///     This Function Reset the Progress Bar's Current Value To Zero
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void ResetProgressBarCore()
        {
            pbrCurrentFile.Value = 0;
        }

        /// <summary>
        ///     This Function Set the Current File's State
        /// </summary>
        /// <param name="state">
        ///     A String Represent the Current File's Processing State
        /// </param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void SetCurrentStateCore(string state)
        {
            lblCurrentState.Text = state;
        }

        /// <summary>
        ///     This Function Set the Over All File State
        /// </summary>
        /// <param name="state">
        ///     A String Represent the Over All File's Processing State
        /// </param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void SetOverAllFileCore(string state)
        {
            lblFileOverAll.Text = state;
        }

        /// <summary>
        ///     This Function Set the To File's State
        /// </summary>
        /// <param name="state">
        ///     A String Represent the To File's Name
        /// </param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void SetToFileCore(string state)
        {
            lblToFileName.Text = state;
        }

        /// <summary>
        ///     This Function Write Log To the UI
        /// </summary>
        /// <param name="log">
        ///     A Log String
        /// </param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void WriteLogCore(string log)
        {
            txtLog.Text += log;
            isDetail = true;
        }

        /// <summary>
        ///     This Function Close the Window
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void CloseFormCore()
        {
            this.Close();
        }

        /// <summary>
        ///     This Function Disable Cancel Button
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-30
        private void DisableCancelButtonCore()
        {
            btnCancel.Enabled = false;
        }

        /// <summary>
        ///     This Function Set Cancel Button To Exit
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-11-26
        private void SetCancelButtonExitCore()
        {
            btnCancel.Text = TranslatorRes.ResourceManager.GetString("CloseButtonTitle");

            btnCancel.Click -= cancelButtonEventHandler;

            btnCancel.Click += new EventHandler(btnClose_Click);

            btnCancel.Enabled = true;
        }

        #endregion

        #region Delegates For Converter

        /// <summary>
        ///     This Function Set the Progress Bar's Current Value
        /// </summary>
        /// <param name="sender">The Sender of the Event</param>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        public void SetProgressBar(object sender, EventArgs e)
        {
            BeginInvoke(setProgressBarDelegateCall);
        }

        /// <summary>
        ///     This Function Set the Current File's State
        /// </summary>
        /// <param name="sender">The Sender of the Event</param>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        public void SetCurrentState(object sender, EventArgs e)
        {
            object[] param = new object[1];
            //param[0] = ((Act.UofTranslator.UofTranslatorLib.UofEventArgs)e).Message;

            //if (((Act.UofTranslator.UofTranslatorLib.UofEventArgs)e).Message != null)
            //{
            //    BeginInvoke(setCurrentStateDelegateCall, param);
            //}
        }

        /// <summary>
        ///     This Function Write Log To the UI
        /// </summary>
        /// <param name="sender">The Sender of the Event</param>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        public void WriteLog(object sender, EventArgs e)
        {
            object[] param = new object[1];
            //param[0] = ((Act.UofTranslator.UofTranslatorLib.UofEventArgs)e).Message;

            //if (((Act.UofTranslator.UofTranslatorLib.UofEventArgs)e).Message != null)
            //{
            //    BeginInvoke(writeLogDelegateCall, param);
            //}
        }

        #endregion

        /// <summary>
        ///     Event Handler for See Log Button
        /// </summary>
        /// <param name="sender">The Event Sender</param>
        /// <param name="e">The Event Parameter</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-24
        private void btnSeeLog_Click(object sender, EventArgs e)
        {
            if (!isViewingLog)
            {
                // Show the Log
                btnSeeLog.Text = btnSeeLog.Text.Replace(">>>", "<<<");
                ShowLog();

                if (lblCurrentState.Text.Equals(TranslatorRes.ResourceManager.GetString("ForDetailPrompt")))
                {
                    lblCurrentState.Text = "";
                }

                // Change the Flag
                isViewingLog = true;
            }
            else
            {
                // Hide the Log
                btnSeeLog.Text = btnSeeLog.Text.Replace("<<<", ">>>");
                HideLog();

                // Change the Flag
                isViewingLog = false;
            }
        }

        /// <summary>
        ///     Event Handler for Closed The Form
        /// </summary>
        /// <param name="sender">The Event Sender</param>
        /// <param name="e">The Event Parameter</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void MainDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                cfgManager.SaveConfig();
            }
            catch (Exception)
            {
                // Do Nothing
            }
            if (isAllSucceed || isOneSucceed)
            {
                // Null
            }
            /*object missing = Type.Missing;
            DialogResult isOpenPartly = DialogResult.Yes;
            if ((!isAllSucceed) && (isOneSucceed) && (TranslatorType == DialogType.Open))
            {
                // Open But not All succeed, let user to choose
                 isOpenPartly = MessageBox.Show(TranslatorRes.ResourceManager.GetString("OpenPartlyPrompt"), 
                                        TranslatorRes.ResourceManager.GetString("OpenPartlyPromptTitle"),
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                 
            }*/
            // All Files Successfully Translated, Open Them 
            // ( Moved to Connect.cs )
            /* if (((isOpenPartly == DialogResult.Yes) || isAllSucceed) && (TranslatorType == DialogType.Open))
            {
                if (wordApplication != null)
                {
                    foreach (FileState currentFileState in FileStates)
                    {
                        if (currentFileState.IsSucceed)
                        {
                            object fileName = currentFileState.DestinationFileName;
                            object readOnly = true;
                            object addToRecentFiles = false;
                            object isVisible = true;
                            object openAndRepair = false;
                            MSword.Document doc = wordApplication.Documents.Open(ref fileName, ref missing, ref readOnly,
                                                                            ref addToRecentFiles, ref missing, ref missing,
                                                                            ref missing, ref missing, ref missing,
                                                                            ref missing, ref missing, ref isVisible,
                                                                            ref openAndRepair, ref missing, ref missing, 
                                                                            ref missing);

                            doc.Fields.Update();
                            doc.Activate();
                        }
                    }
                }
            }*/
        }


        /// <summary>
        ///     Event Handler for Button "Cancel" When It changed to "Close"
        /// </summary>
        /// <param name="sender">The Event Sender</param>
        /// <param name="e">The Event Parameter</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        ///     Event Handler for Button "Cancel"
        /// </summary>
        /// <param name="sender">The Event Sender</param>
        /// <param name="e">The Event Parameter</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void btnCancel_Click(object sender, EventArgs e)
        {
            int SucceedCount = 0;

            Cancel();

            foreach (FileState currentFileState in FileStates)
            {
                if (currentFileState.IsSucceed)
                {
                    SucceedCount++;
                    isOneSucceed = true;
                }
            }

            if (SucceedCount < FileStates.Count)
            {
                isAllSucceed = false;
            }
            else
            {
                isAllSucceed = true;
            }

            this.Close();
        }

        /// <summary>
        ///     Event Handler For Handle Destroyed
        /// </summary>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        protected override void OnHandleDestroyed(EventArgs e)
        {
            // If the handle is being destroyed and you are not
            // recreating it, then abort the translation.

            int SucceedCount = 0;

            if (!RecreatingHandle)
            {
                Cancel();
                foreach (FileState currentFileState in FileStates)
                {
                    if (currentFileState.IsSucceed)
                    {
                        SucceedCount++;
                        isOneSucceed = true;
                    }
                }

                if (SucceedCount < FileStates.Count)
                {
                    isAllSucceed = false;
                }
                else
                {
                    isAllSucceed = true;
                }
            }

            base.OnHandleDestroyed(e);

            if (!RecreatingHandle)
            {
                this.Close();
            }
        }

        /// <summary>
        ///     Event Handler For Handle Created
        /// </summary>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            if (deferTranslate)
            {
                deferTranslate = false;
                StartTranslate();
            }
        }

        /// <summary>
        ///     Event Handler For CheckBox Checked
        /// </summary>
        /// <param name="sender">Event Sender</param>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void chkbxIgnoreError_CheckedChanged(object sender, EventArgs e)
        {
            cfgManager.IsErrorIgnored = chkbxIgnoreError.Checked;
        }

        /// <summary>
        ///     Event Handler For Form Loading
        /// </summary>
        /// <param name="sender">Event Sender</param>
        /// <param name="e">The Argument of the Event</param>
        ///     Author:         µË×·
        ///     Create Date:    2007-05-25
        private void MainDialog_Load(object sender, EventArgs e)
        {
        }

        private void pnlTop_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}