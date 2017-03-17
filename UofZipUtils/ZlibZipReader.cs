/* 
 * Copyright (c) 2006, Clever Age
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
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
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
using System.IO;
using System.Text;
using System.Collections.Generic;


namespace Act.UofTranslator.UofZipUtils
{
    class ZlibZipReader : ZipReader, IDisposable
    {
        private IntPtr handle;
        private bool disposed = false;

        internal ZlibZipReader(string path)
        {
            string resolvedPath = ZipLib.ResolvePath(path);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException("File does not exist:" + path);
            }
            this.handle = ZipLib.unzOpen(resolvedPath);
            if (handle == IntPtr.Zero)
            {
                throw new ZipException("Unable to open ZIP file:" + path);
            }
        }

        ~ZlibZipReader()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources : none here
                }
                Close();
            }
        }

        public override void Close()
        {
            if (handle != IntPtr.Zero)
            {
                int result = ZipLib.unzClose(this.handle);
                handle = IntPtr.Zero;
                // Question: should we raise this exception ?
                if (result != 0)
                {
                    throw new ZipException("Error closing file - Errorcode: " + result);
                }
            }
        }

        public override Stream GetEntry(string relativePath)
        {
            string resolvedPath = ZipLib.ResolvePath(relativePath);
            if (ZipLib.unzLocateFile(this.handle, resolvedPath, 0) != 0)
            {
                throw new ZipEntryNotFoundException("Entry not found:" + relativePath);
            }

            ZipEntryInfo entryInfo = new ZipEntryInfo();
            int result = ZipLib.unzGetCurrentFileInfo(this.handle, out entryInfo, null, 0, null, 0, null, 0);
            if (result != 0)
            {
                throw new ZipException("Error while reading entry info: " + relativePath + " - Errorcode: " + result);
            }

            result = ZipLib.unzOpenCurrentFile(this.handle);
            if (result != 0)
            {
                throw new ZipException("Error while opening entry: " + relativePath + " - Errorcode: " + result);
            }

            byte[] buffer = new byte[entryInfo.UncompressedSize];
            int bytesRead = 0;
            if ((bytesRead = ZipLib.unzReadCurrentFile(this.handle, buffer, (uint)entryInfo.UncompressedSize)) < 0)
            {
                throw new ZipException("Error while reading entry: " + relativePath + " - Errorcode: " + result);
            }

            result = ZipLib.unzCloseCurrentFile(handle);
            if (result != 0)
            {
                throw new ZipException("Error while closing entry: " + relativePath + " - Errorcode: " + result);
            }

            return new MemoryStream(buffer, 0, bytesRead);
        }

        public override void ExtractOfficeDocument(string zipFileName, string directory)
        {
            IntPtr zhandle = ZipLib.unzOpen(zipFileName);
            FileStream fs = new FileStream(zipFileName, FileMode.Open, FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            int i = 0;
            int flag = 0;
            int flag2 = 0;
            int count = 1;
            try
            {
                while (br.BaseStream.Position < br.BaseStream.Length)
                {
                    if (br.ReadByte() == 0x50)
                    {
                        if (br.ReadByte() == 0x4b)
                        {
                            if (br.ReadByte() == 0x03)
                            {
                                if (br.ReadByte() == 0x04)
                                {

                                    br.BaseStream.Position += 22;
                                    int filenameLength = br.ReadInt16();
                                    br.BaseStream.Position += 2;
                                    byte[] buffer = new byte[filenameLength];
                                    br.BaseStream.Read(buffer, 0, filenameLength);
                                    br.BaseStream.Position -= (26 + filenameLength);
                                    string fullFileName = directory + ZipLib.ResolvePath(Encoding.GetEncoding("UTF-8").GetString(buffer));
                                    DirectoryInfo dir = new DirectoryInfo(fullFileName);
                                    if (fullFileName.Contains("/"))
                                    {
                                        string tempdir = "";
                                        while (tempdir != dir.Root.FullName)
                                        {
                                            tempdir = dir.Parent.FullName;
                                            dir = new DirectoryInfo(dir.Parent.FullName);
                                            if (!Directory.Exists(tempdir))
                                            {
                                                Directory.CreateDirectory(tempdir);
                                            }
                                        }
                                    }

                                    FileStream writeFile = File.Create(fullFileName);
                                    if (i < count)
                                    {
                                        try
                                        {
                                            i++;
                                            if (flag == 0)
                                            {
                                                ZipLib.unzGoToFirstFile(zhandle);
                                                flag = 1;
                                            }
                                            else
                                            {
                                                int fir = ZipLib.unzGoToNextFile(zhandle);
                                            }
                                            ZipEntryInfo zipentryInfo;
                                            ZipLib.unzGetCurrentFileInfo(zhandle, out zipentryInfo, null, 0, null, 0, null, 0);
                                            ZipLib.unzOpenCurrentFile(zhandle);
                                            ZipFileInfo zipfileinfo;
                                            ZipLib.unzGetGlobalInfo(zhandle, out zipfileinfo);
                                            if (flag2 == 0)
                                            {
                                                count = (int)zipfileinfo.EntryCount;
                                                flag2 = 1;
                                            }
                                            byte[] buffer2 = new byte[zipentryInfo.UncompressedSize];
                                            ZipLib.unzReadCurrentFile(zhandle, buffer2, (uint)zipentryInfo.UncompressedSize);
                                            writeFile.Write(buffer2, 0, (int)zipentryInfo.UncompressedSize);
                                            br.BaseStream.Position += zipentryInfo.CompressedSize;
                                        }
                                        catch (Exception ex)
                                        {
                                            throw ex;
                                        }
                                        finally
                                        {
                                            if (writeFile != null)
                                                writeFile.Close();
                                            if (ZipLib.unzCloseCurrentFile(zhandle) != 0)
                                                ZipLib.unzCloseCurrentFile(zhandle);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
                if (br != null)
                    br.Close();
            }

        }

        public override void ExtractUOFDocument(string fileName, string directory)
        {
            List<string> dirInfo = GetZipDirInfo(fileName);
            foreach (string info in dirInfo)
            {
                if (info.EndsWith("/"))
                {
                    CreateDirectory(directory, info.Substring(0, info.Length - 1));
                }
                else
                {
                    if (info.Contains("/")||info.Contains("\\"))
                    {
                        string subDir = info;
                        if (subDir.Contains("\\"))
                        {
                            subDir = subDir.Replace("\\", "/");
                        }

                        while (subDir.Contains("/") )
                        {
                            
                            CreateDirectory(directory, subDir.Substring(0, subDir.IndexOf("/")));
                            subDir = subDir.Substring(subDir.IndexOf("/") + 1, subDir.Length - subDir.IndexOf("/") - 1);
                        }

                    }
                    BinaryReader br = null;
                    BinaryWriter bw = null;
                    try
                    {
                        ZipReader archive = ZipFactory.OpenArchive(fileName);
                        string tmpFileName = info.Replace('/', Path.AltDirectorySeparatorChar);

                        br = new BinaryReader(archive.GetEntry(info));
                        bw = new BinaryWriter(File.Open(directory + Path.AltDirectorySeparatorChar + tmpFileName, FileMode.Create));
                        byte[] buffer = new byte[br.BaseStream.Length];

                        br.Read(buffer, 0, (int)br.BaseStream.Length);
                        bw.Write(buffer);
                        bw.Flush();
                        bw.Close();
                        br.Close();
                        archive.Close();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.StackTrace);
                    }
                    finally
                    {
                        if (br != null)
                        {
                            br.Close();
                        }
                        if (bw != null)
                        {
                            bw.Close();
                        }
                    }
                }
            }
        }

        public override List<string> GetZipDirInfo(string zipFileName)
        {
            //IntPtr zhandle = ZipLib.unzOpen(zipFileName);
            FileStream fs = new FileStream(zipFileName, FileMode.Open, FileAccess.Read);
            BinaryReader br = new BinaryReader(fs);
            List<string> zipDirInfo = new List<string>();

            try
            {
                while (br.BaseStream.Position < br.BaseStream.Length)
                {
                    if (br.ReadByte() == 0x50)
                    {
                        if (br.ReadByte() == 0x4b)
                        {
                            if (br.ReadByte() == 0x01)
                            {
                                if (br.ReadByte() == 0x02)
                                {
                                    // 22=2version made by+2version needed to extract+2general purpose bit flag+2compression method+2last mod file time+
                                    // 2last mode file date+4crc-32+4compressedsize+4uncompressedsize
                                    br.BaseStream.Position += 24;
                                    int filenameLength = br.ReadInt16();
                                    // br.BaseStream.Position += 2;

                                    // 16=2extra field length+2file comment length+2disk number start+4external file attributes+4relative offset of local header
                                    br.BaseStream.Position += 16;

                                    byte[] buffer = new byte[filenameLength];
                                    br.BaseStream.Read(buffer, 0, filenameLength);

                                    string fileName = ZipLib.ResolvePath(Encoding.GetEncoding("UTF-8").GetString(buffer));
                                    zipDirInfo.Add(fileName);

                                }
                            }
                        }
                    }
                }

                return zipDirInfo;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
                if (br != null)
                    br.Close();
            }

        }

        private void CreateDirectory(string directory, string subdirectory)
        {
            if (!Directory.Exists(directory + Path.AltDirectorySeparatorChar + subdirectory))
            {
                Directory.CreateDirectory(directory + Path.AltDirectorySeparatorChar + subdirectory);
            }
        }
    }
}
