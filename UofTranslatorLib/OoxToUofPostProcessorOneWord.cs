using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Act.UofTranslator.UofZipUtils;
using System.IO;
using System.Xml.XPath;
using System.Xml.Xsl;
using log4net;
using System.Web;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// The first step post process
    /// </summary>
    /// <author>fangchunyan</author>
    class OoxToUofPostProcessorOneWord : AbstractProcessor
    {

        private XmlNamespaceManager namespaceManager;

        private static ILog logger = LogManager.GetLogger(typeof(OoxToUofPostProcessorOneWord).FullName);

        public override bool transform()
        {
            bool result = true;
            XmlTextWriter resultWriter = null;

            //string posttempfile1 = Path.GetTempPath() + "posttempfile1.xml";
            string posttempfile1 = Path.GetDirectoryName(originalFile) + "\\" + "posttempfile1.xml";
            try
            {
                XmlDocument xmlDoc = new XmlDocument();//xml�ĵ�
                xmlDoc.Load(inputFile);//��ָ��url����xml�ĵ���������ת���õ����м��ĵ�

                XmlNameTable table = xmlDoc.NameTable;
                namespaceManager = new XmlNamespaceManager(table);
                namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                namespaceManager.AddNamespace("��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                namespaceManager.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
                namespaceManager.AddNamespace("ͼ", "http://schemas.uof.org/cn/2003/graph");

                XmlNode root = xmlDoc.DocumentElement;
                RestructureBookmark(xmlDoc, root);
                RestructureComments(xmlDoc, root);
                RestructureHyperlink(xmlDoc, root);
                RemoveDoubleFont(xmlDoc, root);
                RestructureInstrText(xmlDoc, root);

                resultWriter = new XmlTextWriter(posttempfile1, System.Text.Encoding.UTF8);
                xmlDoc.Save(resultWriter);
                resultWriter.Close();
                DeleteFirstChars(posttempfile1, outputFile);

            }
            catch (Exception e)
            {
                result = false;
                logger.Error(e.Message);
                logger.Error(e.StackTrace);
            }
            finally
            {
                if (resultWriter != null)
                    resultWriter.Close();
            }
            return result;
        }


        private void DeleteFirstChars(string source, string result)//ɾ��xml:space=\"preserve\"
        {
            StreamReader sr = new StreamReader(source);
            StreamWriter sw = new StreamWriter(result);
            string strTemp = "";

            while ((strTemp = sr.ReadLine()) != null)
            {
                if (strTemp.Contains("<��:�ı���") == true && strTemp.Contains("xml:space=\"preserve\"") == true)
                {
                    strTemp = strTemp.Replace("xml:space=\"preserve\"", "");

                }
                sw.WriteLine(strTemp);
            }
            sr.Close();
            sw.Close();
        }

        private void RestructureBookmark(XmlDocument xmlDoc, XmlNode root)//�����ǩ
        {
            /*
            XmlNodeList bookmarkList = root.SelectNodes("//w:bookmarkStart", namespaceManager);
            if (bookmarkList.Count != 0)
            {
                XmlElement bookmarkset = xmlDoc.CreateElement("uof", "��ǩ��", "http://schemas.uof.org/cn/2003/uof");
                bookmarkset.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0027");

                for (int i = 0; i < bookmarkList.Count; i++)
                {
                    XmlElement bookmarkStart = (XmlElement)bookmarkList[i];
                    RebuildTblBK(bookmarkStart, root, xmlDoc);

                    XmlElement bookmark = xmlDoc.CreateElement("uof", "��ǩ", "http://schemas.uof.org/cn/2003/uof");
                    bookmark.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0028");
                    bookmark.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "����");
                    string name;
                    if (bookmarkStart.Attributes["w:name"] != null)
                    {
                        name = bookmarkStart.Attributes["w:name"].Value;
                    }
                    else
                    {
                        name = "bookmark" + i;
                    }
                    bookmark.SetAttribute("����", "http://schemas.uof.org/cn/2003/uof", name);

                    XmlElement bktextpos = xmlDoc.CreateElement("uof", "�ı�λ��", "http://schemas.uof.org/cn/2003/uof");
                    bktextpos.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0029");
                    bktextpos.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��������");
                    string regionref = bookmarkStart.Attributes["w:id"].Value;
                    bktextpos.SetAttribute("��������", "http://schemas.uof.org/cn/2003/uof-wordproc", regionref);

                    bookmark.AppendChild(bktextpos);
                    bookmarkset.AppendChild(bookmark);

                    XmlElement bkregionStart = xmlDoc.CreateElement("��", "����ʼ", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    bkregionStart.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0121");
                    bkregionStart.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ�� ���� ����");
                    bkregionStart.SetAttribute("����", "http://schemas.uof.org/cn/2003/uof-wordproc", "bookmark");
                    bkregionStart.SetAttribute("����", "http://schemas.uof.org/cn/2003/uof-wordproc", name);
                    bkregionStart.SetAttribute("��ʶ��", "http://schemas.uof.org/cn/2003/uof-wordproc", regionref);

                    XmlElement bkParent = (XmlElement)bookmarkStart.ParentNode;
                    if (bkParent != null)
                    {
                        bkParent.ReplaceChild(bkregionStart, bookmarkStart);
                    }

                    XmlElement bkend = (XmlElement)root.SelectSingleNode("//w:bookmarkEnd[@w:id='" + regionref + "']", namespaceManager);
                   //zhaobj�޸�
                    if (bkend != null)
                    {
                        XmlElement bkendParent = (XmlElement)bkend.ParentNode;
                        XmlElement bkregionEnd = xmlDoc.CreateElement("��", "�������", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        bkregionEnd.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0122");
                        bkregionEnd.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ������");
                        bkregionEnd.SetAttribute("��ʶ������", "http://schemas.uof.org/cn/2003/uof-wordproc", regionref);

                        bkendParent.ReplaceChild(bkregionEnd, bkend);
                    }

                }
                root.InsertAfter(bookmarkset, root.FirstChild);


            }
            */
        }

        private void RestructureComments(XmlDocument xmlDoc, XmlNode root)
        {
            XmlNodeList commentList = root.SelectNodes("//��:����ʼ[@��:����='annotation']", namespaceManager);
            if (commentList.Count != 0)
            {
                for (int i = 0; i < commentList.Count; i++)
                {
                    XmlElement commentStart = (XmlElement)commentList[i];
                    string id = commentStart.Attributes["��:��ʶ��"].Value;

                    XmlElement commentEnd = (XmlElement)xmlDoc.SelectSingleNode("//��:�������[@��:��ʶ������='" + id + "']", namespaceManager);
                    if (commentEnd != null && commentEnd.ParentNode.Name == "��:����")
                    {
                        XmlElement commentPresibling = (XmlElement)commentEnd.PreviousSibling;
                        XmlElement run = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        run.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                        run.AppendChild(commentEnd);

                        if (commentPresibling.Name == "��:����")
                        {
                            commentPresibling.AppendChild(run);
                        }
                        else if (commentPresibling.Name == "��:���ֱ�")
                        {

                            XmlElement row = (XmlElement)commentPresibling.LastChild;
                            if (row == null || row.Name != "��:��")
                            {
                                XmlElement newrow = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                newrow.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0143");
                                XmlElement cell = xmlDoc.CreateElement("��", "��Ԫ��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                cell.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0148");
                                XmlElement para = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                para.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                                para.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                                para.AppendChild(run);
                                cell.AppendChild(para);
                                commentPresibling.AppendChild(newrow);

                            }
                            else if (row.Name == "��:��")
                            {
                                XmlElement cell = (XmlElement)row.LastChild;
                                if (cell == null && cell.Name != "��:��Ԫ��")
                                {
                                    XmlElement newcell = xmlDoc.CreateElement("��", "��Ԫ��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                    newcell.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0148");
                                    XmlElement para = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                    para.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                                    para.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                                    para.AppendChild(run);
                                    cell.AppendChild(para);
                                    row.AppendChild(newcell);
                                }
                                else if (cell.Name == "��:��Ԫ��")
                                {
                                    XmlElement para = (XmlElement)cell.LastChild;
                                    if (para == null && para.Name != "��:����")
                                    {
                                        XmlElement newpara = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                        newpara.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                                        newpara.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                                        newpara.AppendChild(run);
                                        cell.AppendChild(para);
                                    }
                                    else
                                    {
                                        para.AppendChild(run);
                                    }
                                }
                            }
                        }
                        else
                        {
                            XmlElement para = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                            para.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                            para.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                            para.AppendChild(run);
                            XmlElement body = (XmlElement)commentPresibling.ParentNode;
                            body.InsertAfter(para, commentPresibling);
                        }
                    }
                    else if (commentEnd != null && commentEnd.ParentNode.Name == "����")
                    {
                        XmlElement commentPresibling = (XmlElement)commentEnd.PreviousSibling;
                        XmlElement run = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        run.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                        run.AppendChild(commentEnd);

                        commentEnd.ParentNode.InsertAfter(run, commentPresibling);

                    }




                }

            }
        }



        private void RebuildTblBK(XmlElement bkstart, XmlNode root, XmlDocument xmlDoc)
        {
            string id = bkstart.Attributes["w:id"].Value;
            XmlElement bkend = (XmlElement)root.SelectSingleNode("//w:bookmarkEnd[@w:id='" + id + "']", namespaceManager);
            if (bkend == null)
            {
                return;
            }
            XmlElement bkEndParent = (XmlElement)bkend.ParentNode;
            if (bkEndParent.Name == "��:���ֱ�")
            {
                dealTblBookmark(bkEndParent, bkstart, bkend, xmlDoc);
            }
            else if (bkEndParent.Name == "��:����")
            {
                XmlElement bkPresibling = (XmlElement)bkend.PreviousSibling;
                XmlElement run = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                run.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                run.AppendChild(bkend);

                if (bkPresibling.Name == "��:����")
                {
                    bkPresibling.AppendChild(run);

                }
                else if (bkPresibling.Name == "��:���ֱ�")
                {

                    XmlElement row = (XmlElement)bkPresibling.LastChild;
                    if (row == null || row.Name != "��:��")
                    {
                        XmlElement newrow = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        newrow.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0143");
                        XmlElement cell = xmlDoc.CreateElement("��", "��Ԫ��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        cell.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0148");
                        XmlElement para = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        para.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                        para.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                        para.AppendChild(run);
                        cell.AppendChild(para);
                        bkPresibling.AppendChild(newrow);

                    }
                    else if (row.Name == "��:��")
                    {
                        XmlElement cell = (XmlElement)row.LastChild;
                        if (cell == null && cell.Name != "��:��Ԫ��")
                        {
                            XmlElement newcell = xmlDoc.CreateElement("��", "��Ԫ��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                            newcell.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0148");
                            XmlElement para = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                            para.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                            para.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                            para.AppendChild(run);
                            cell.AppendChild(para);
                            row.AppendChild(newcell);
                        }
                        else if (cell.Name == "��:��Ԫ��")
                        {
                            XmlElement para = (XmlElement)cell.LastChild;
                            if (para == null && para.Name != "��:����")
                            {
                                XmlElement newpara = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                newpara.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                                newpara.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                                newpara.AppendChild(run);
                                cell.AppendChild(para);
                            }
                            else
                            {
                                para.AppendChild(run);
                            }
                        }
                    }
                }
                else
                {
                    XmlElement para = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    para.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                    para.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                    para.AppendChild(run);
                    XmlElement body = (XmlElement)bkPresibling.ParentNode;
                    body.InsertAfter(para, bkPresibling);
                }
            }
            else if (bkEndParent.Name == "��:����")
            {
                XmlElement bkPresibling = (XmlElement)bkend.PreviousSibling;
                XmlElement run = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                run.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                run.AppendChild(bkend);
                bkEndParent.InsertAfter(run, bkPresibling);


            }
            return;
        }

        private void dealTblBookmark(XmlElement tbl, XmlElement bkstart, XmlElement bkend, XmlDocument xmlDoc)
        {
            string id = bkstart.Attributes["w:id"].Value;

            if (bkstart.HasAttribute("w:colFirst") || bkstart.HasAttribute("w:colLast"))
            {
                int bkstartcolNum = Convert.ToInt32(bkstart.Attributes["w:colFirst"].Value);
                XmlElement para = (XmlElement)bkstart.ParentNode;
                XmlElement tc = (XmlElement)para.ParentNode;
                XmlElement tr = (XmlElement)tc.ParentNode;

                string code = "��:��Ԫ��[" + bkstartcolNum + "]";
                XmlElement currtc = (XmlElement)tr.SelectSingleNode(code, namespaceManager);
                if (currtc == null)
                {
                    XmlElement newcell = xmlDoc.CreateElement("��", "��Ԫ��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    newcell.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0148");
                    XmlElement newpara = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    newpara.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                    newpara.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                    XmlElement newrun = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    newrun.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                    newrun.AppendChild(bkstart);
                    newpara.AppendChild(newrun);
                    newcell.AppendChild(newpara);

                }
                else
                {
                    XmlElement currpara = (XmlElement)currtc.SelectSingleNode("��:����[last()]", namespaceManager);
                    if (currpara == null)
                    {
                        XmlElement newpara = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        newpara.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                        newpara.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");
                        XmlElement newrun = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                        newrun.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                        newrun.AppendChild(bkstart);
                        newpara.AppendChild(newrun);
                        currtc.AppendChild(newpara);
                    }
                    else
                    {
                        XmlElement currrun = (XmlElement)currpara.SelectSingleNode("��:��[last()]", namespaceManager);
                        if (currrun == null)
                        {
                            XmlElement newrun = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                            newrun.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                            newrun.AppendChild(bkstart);
                            currpara.AppendChild(newrun);
                            currtc.AppendChild(currpara);
                        }
                        else
                        {
                            currrun.AppendChild(bkstart);
                        }
                    }
                }

                XmlElement rowBeforebkEnd = (XmlElement)bkend.PreviousSibling;
                int bkendcolNum = Convert.ToInt32(bkstart.Attributes["w:colLast"].Value) + 1;
                XmlElement cell = (XmlElement)rowBeforebkEnd.SelectSingleNode("��:��Ԫ��[" + bkendcolNum + "]", namespaceManager);
                XmlElement lastPara = (XmlElement)cell.SelectSingleNode("��:����[last()]", namespaceManager);

                XmlElement bkregionEndRun = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                bkregionEndRun.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                bkregionEndRun.AppendChild(bkend);
                lastPara.AppendChild(bkregionEndRun);


            }
            else
            {
                XmlElement rowBeforebkEnd = (XmlElement)bkend.PreviousSibling;
                XmlElement lastCell = (XmlElement)rowBeforebkEnd.LastChild;//maybe error
                XmlElement lastPara = (XmlElement)lastCell.LastChild;

                XmlElement bkregionEndRun = xmlDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
                bkregionEndRun.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0085");
                bkregionEndRun.AppendChild(bkend);
                lastPara.AppendChild(bkregionEndRun);


            }

            return;


        }


        private void RestructureHyperlink(XmlDocument xmlDoc, XmlNode root)
        {
            XmlNodeList hyperlinkStartList = root.SelectNodes("//��:����ʼ[@��:����='hyperlink']", namespaceManager);
            if (hyperlinkStartList.Count != 0)
            {
                XmlElement hyperlinkset = xmlDoc.CreateElement("uof", "���Ӽ�", "http://schemas.uof.org/cn/2003/uof");
                hyperlinkset.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0031");

                for (int i = 0; i < hyperlinkStartList.Count; i++)
                {
                    XmlElement hyperlinkStart = (XmlElement)hyperlinkStartList[i];

                    XmlElement hyperlink = xmlDoc.CreateElement("uof", "��������", "http://schemas.uof.org/cn/2003/uof");
                    hyperlink.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0032");
                    hyperlink.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ�� Ŀ�� ��ǩ ʽ������ �ѷ���ʽ������ ��ʾ ��Դ");
                    string hlid = "hlid_" + i;
                    hyperlink.SetAttribute("��ʶ��", "http://schemas.uof.org/cn/2003/uof", hlid);

                    if (hyperlinkStart.Attributes["uof:Ŀ��"] != null)
                    {
                        string target = hyperlinkStart.Attributes["uof:Ŀ��"].Value;
                        hyperlink.SetAttribute("Ŀ��", "http://schemas.uof.org/cn/2003/uof", HttpUtility.UrlDecode(target));
                        hyperlinkStart.RemoveAttribute("Ŀ��", "http://schemas.uof.org/cn/2003/uof");
                    }
                    if (hyperlinkStart.Attributes["uof:��ǩ"] != null)
                    {
                        string bkid = hyperlinkStart.Attributes["uof:��ǩ"].Value;
                        hyperlink.SetAttribute("��ǩ", "http://schemas.uof.org/cn/2003/uof", bkid);
                        hyperlinkStart.RemoveAttribute("��ǩ", "http://schemas.uof.org/cn/2003/uof");
                    }

                    if (hyperlinkStart.Attributes["uof:��ʾ"] != null)
                    {
                        string tip = hyperlinkStart.Attributes["uof:��ʾ"].Value;
                        hyperlink.SetAttribute("��ʾ", "http://schemas.uof.org/cn/2003/uof", tip);
                        hyperlinkStart.RemoveAttribute("��ʾ", "http://schemas.uof.org/cn/2003/uof");
                    }
                    if (hyperlinkStart.Attributes["uof:��Դ"] != null)
                    {
                        string source = hyperlinkStart.Attributes["uof:��Դ"].Value;
                        hyperlink.SetAttribute("��Դ", "http://schemas.uof.org/cn/2003/uof", source);
                        hyperlinkStart.RemoveAttribute("��Դ", "http://schemas.uof.org/cn/2003/uof");
                    }

                    hyperlinkset.AppendChild(hyperlink);
                }

                //XmlElement bkset = (XmlElement)root.SelectSingleNode("//uof:��ǩ��", namespaceManager);
                //if (bkset != null)
                //{
                //    root.InsertAfter(hyperlinkset, bkset);
                //}
                //else
                //{
                //    root.InsertAfter(hyperlinkset, root.FirstChild);
                //}


            }
        }


        private void RemoveDoubleFont(XmlDocument xmlDoc, XmlNode root)
        {
            XmlElement defaultRunStyle = (XmlElement)root.SelectSingleNode("//uof:��ʽ��[@��:����='default']", namespaceManager);
            if (defaultRunStyle != null)
            {
                RemoveDoubleRfont(defaultRunStyle, "��:����");
            }

            XmlElement defaultRunProperty = (XmlElement)root.SelectSingleNode("//uof:����ʽ��[@��:����='default']/��:������", namespaceManager);
            if (defaultRunProperty != null)
            {
                RemoveDoubleRfont(defaultRunProperty, "��:����");
            }

            XmlElement defaultParaStyle = (XmlElement)root.SelectSingleNode("//uof:����ʽ��[@��:����='default']", namespaceManager);
            if (defaultParaStyle != null)
            {
                RemoveDoubleRfont(defaultParaStyle, "��:����");
            }

        }

        private void RemoveDoubleRfont(XmlElement defaultRunStyle, string nodename)
        {
            XmlNodeList fonts = defaultRunStyle.SelectNodes(nodename, namespaceManager);
            if (fonts.Count > 1)
            {
                XmlElement firstFont = (XmlElement)fonts[0];
                XmlElement otherFont = (XmlElement)fonts[1];
                for (int j = 0; j < otherFont.Attributes.Count; j++)
                {
                    XmlAttribute otherFontAttr = otherFont.Attributes[j];
                    string otherAttrName = otherFontAttr.Name;
                    if (firstFont.Attributes[otherAttrName] == null)
                    {
                        string otherAttrValue = otherFontAttr.Value;
                        firstFont.SetAttribute(otherAttrName, otherAttrValue);
                    }
                }

                XmlNode fontParent = otherFont.ParentNode;
                fontParent.RemoveChild(otherFont);


            }

        }

        private void RestructureInstrText(XmlDocument xmlDoc, XmlNode root)
        {
            XmlNodeList InstrTextStartList = root.SelectNodes("//w:fldChar[@w:fldCharType='begin']", namespaceManager);
            for (int i = 0; i < InstrTextStartList.Count; i++)
            {
                XmlNode fldCharStart = InstrTextStartList[i];
                XmlElement fldCharStartRun = (XmlElement)fldCharStart.ParentNode;
                XmlElement fldCharStartPara = (XmlElement)fldCharStartRun.ParentNode;
                XmlNodeList instrList = fldCharStartRun.SelectNodes("following-sibling::node()[following-sibling::node()/w:fldChar[1]/@w:fldCharType='separate']", namespaceManager);
                string code = "";
                for (int j = 0; j < instrList.Count; j++)
                {
                    XmlNodeList instrCodeList = instrList[j].SelectNodes("w:instrText", namespaceManager);
                    for (int m = 0; m < instrCodeList.Count; m++)
                        code = code + instrCodeList[m].InnerText;

                }
                if (code.Contains("AUTHOR") || code.Contains("FILENAME") || code.Contains("REF") || code.Contains("SEQ")
                    || code.Contains("XE") || code.Contains("TIME") || code.Contains("PAGE") || code.Contains("TITLE")
                    || code.Contains("SECTION") || code.Contains("NUMCHARS") || code.Contains("SAVEDATE")
                     || code.Contains("CREATEDATE") || code.Contains("NUMPAGES") || code.Contains("REVNUM")
                     || code.Contains("SECTIONPAGES"))
                {
                    string codetype = RegionCodeTable(code);
                    XmlElement regionStart = xmlDoc.CreateElement("��", "��ʼ", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    regionStart.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0079");
                    regionStart.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "���� ����");
                    regionStart.SetAttribute("����", "http://schemas.uof.org/cn/2003/uof-wordproc", codetype);
                    regionStart.SetAttribute("����", "http://schemas.uof.org/cn/2003/uof-wordproc", "false");
                    XmlElement regionCode = xmlDoc.CreateElement("��", "�����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    regionCode.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0080");
                    XmlElement regionPara = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    regionPara.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0051");
                    regionPara.SetAttribute("attrList", "http://schemas.uof.org/cn/2003/uof", "��ʶ��");

                    XmlElement StartRunNext = (XmlElement)fldCharStartRun.NextSibling;

                    while (StartRunNext != null && StartRunNext.SelectSingleNode("w:fldChar", namespaceManager) == null)
                    {
                        StartRunNext = (XmlElement)StartRunNext.NextSibling;
                        regionPara.AppendChild(StartRunNext.PreviousSibling);
                    }

                    regionCode.AppendChild(regionPara);

                    fldCharStartPara.InsertAfter(regionCode, fldCharStartRun);
                    fldCharStartPara.InsertAfter(regionStart, fldCharStartRun);

                    XmlNode instrEndRun = fldCharStartRun.SelectSingleNode("following-sibling::node()[w:fldChar/@w:fldCharType='end'][1]", namespaceManager);
                    XmlNode instrEnd = instrEndRun.SelectSingleNode("w:fldChar[1][@w:fldCharType='end']", namespaceManager);
                    XmlNode regionEndPara = instrEndRun.ParentNode;
                    XmlElement regionEnd = xmlDoc.CreateElement("��", "�����", "http://schemas.uof.org/cn/2003/uof-wordproc");
                    regionEnd.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "t0081");
                    regionEndPara.ReplaceChild(regionEnd, instrEndRun);

                }


            }


            XmlNodeList list = root.SelectNodes("//w:fldChar", namespaceManager);
            for (int i = 0; i < list.Count; i++)
            {
                XmlNode node = list[i];
                XmlNode regionEndRun = node.ParentNode;
                regionEndRun.RemoveChild(node);
            }

            XmlNodeList instrTextList = root.SelectNodes("//w:instrText", namespaceManager);
            for (int i = 0; i < instrTextList.Count; i++)
            {
                XmlNode node = instrTextList[i];
                XmlNode regionEndRun = node.ParentNode;
                regionEndRun.RemoveChild(node);
            }

        }


        private string RegionCodeTable(string code)
        {
            string tmp = code.ToLower().Trim();
            string result = "";
            if (tmp.Contains("revision"))
            {
                result = "revnum";
            }
            else if (tmp.Contains("sectionpages"))
            {
                result = "pageinsection";
            }
            else if (tmp.Contains("author"))
            {
                result = "author";
            }
            else if (tmp.Contains("filename"))
            {
                result = "filename";
            }
            else if (tmp.Contains("ref"))
            {
                result = "ref";
            }
            else if (tmp.Contains("xe"))
            {
                result = "xe";
            }
            else if (tmp.Contains("seq"))
            {
                result = "seq";
            }
            else if (tmp.Contains("link"))
            {
                result = "link";
            }
            else if (tmp.Contains("time"))
            {
                result = "time";
            }
            else if (tmp.Contains("page"))
            {
                result = "page";
            }
            else if (tmp.Contains("title"))
            {
                result = "title";
            }
            else if (tmp.Contains("section"))
            {
                result = "section";
            }
            else if (tmp.Contains("numchars"))
            {
                result = "numchars";
            }
            else if (tmp.Contains("savedate"))
            {
                result = "savedate";

            }
            else if (tmp.Contains("createdate"))
            {
                result = "createdate";

            }
            else if (tmp.Contains("numpages"))
            {
                result = "numpages";

            }
            return result;

        }


    }
}