using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security.AccessControl;

namespace CustomSetupActions
{
    /// <summary>
    ///  register uoftranslator in x64 environment
    /// </summary>
    public class OfficeRegister
    {
        
        [CustomAction]
        public static ActionResult X64OfficeRegister(Session session)
        {
            session.Log("Begin register UOFTranslator X64 Office");
            Record record = new Record(0);

            try
            {
                if (Environment.Is64BitOperatingSystem)
                {

                    RegistryKey registryBase = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
                   // System.Windows.Forms.MessageBox.Show("open base key" + registryBase.ValueCount);
                    //NTAccount name = new NTAccount(Environment.UserName);
                    //RegistrySecurity rs = Registry.LocalMachine.OpenSubKey("software").OpenSubKey("Microsoft").GetAccessControl();
                    
                    //rs.AddAccessRule(new RegistryAccessRule(name, RegistryRights.FullControl, AccessControlType.Allow));
                    //Registry.LocalMachine.OpenSubKey("software").OpenSubKey("Microsoft", true).SetAccessControl(rs);

                    
                    //RegistryKey officeRegistryBase = Registry.LocalMachine.OpenSubKey("software", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl).OpenSubKey("Microsoft", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.FullControl);
                    string regpath = @"SOFTWARE\Microsoft\";
                    RegistryKey officeRegistryBase = registryBase.OpenSubKey(regpath,
                        RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);
                   // System.Windows.Forms.MessageBox.Show("open microsoft key" + officeRegistryBase.ValueCount);

                   // System.Security.AccessControl.RegistrySecurity sec=new System.Security.AccessControl.RegistrySecurity();

                    if (officeRegistryBase != null)
                    {
                        // register powerpoint x64
                        RegistryKey powerpntExport = officeRegistryBase.CreateSubKey(@"Office\15.0\PowerPoint\Presentation Converters\OOXML Converters\Export\UOF Presentation Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        powerpntExport.SetValue("Clsid", "{AEC7E6B1-3D6E-41E0-A442-34A0D6995DA4}", RegistryValueKind.String);
                        powerpntExport.SetValue("Extensions", "uop", RegistryValueKind.String);
                        powerpntExport.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        RegistryKey powerpntImport = officeRegistryBase.CreateSubKey(@"Office\15.0\PowerPoint\Presentation Converters\OOXML Converters\Import\UOF Presentation Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        powerpntImport.SetValue("Clsid", "{AEC7E6B1-3D6E-41E0-A442-34A0D6995DA4}", RegistryValueKind.String);
                        powerpntImport.SetValue("Extensions", "uop", RegistryValueKind.String);
                        powerpntImport.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        // register word x64
                        RegistryKey wordExport = officeRegistryBase.CreateSubKey(@"Office\15.0\Word\Text Converters\OOXML Converters\Export\UOF Word Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        wordExport.SetValue("Clsid", "{6970F0A4-88D6-4BAE-BC69-5C40E3598336}", RegistryValueKind.String);
                        wordExport.SetValue("Extensions", "uot", RegistryValueKind.String);
                        wordExport.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        RegistryKey wordImport = officeRegistryBase.CreateSubKey(@"Office\15.0\Word\Text Converters\OOXML Converters\Import\UOF Word Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        wordImport.SetValue("Clsid", "{6970F0A4-88D6-4BAE-BC69-5C40E3598336}", RegistryValueKind.String);
                        wordImport.SetValue("Extensions", "uot", RegistryValueKind.String);
                        wordImport.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        // register excel x64
                        RegistryKey excelExport = officeRegistryBase.CreateSubKey(@"Office\15.0\Excel\Text Converters\OOXML Converters\Export\UOF Spreadsheet Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        excelExport.SetValue("Clsid", "{ADBE850B-D37F-4422-B66E-88471BDC1B20}", RegistryValueKind.String);
                        excelExport.SetValue("Extensions", "uos", RegistryValueKind.String);
                        excelExport.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        RegistryKey excelImport = officeRegistryBase.CreateSubKey(@"Office\15.0\Excel\Text Converters\OOXML ConvertersImport\UOF Spreadsheet Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        excelImport.SetValue("Clsid", "{ADBE850B-D37F-4422-B66E-88471BDC1B20}", RegistryValueKind.String);
                        excelImport.SetValue("Extensions", "uos", RegistryValueKind.String);
                        excelImport.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        // support office 2010 x64
                        // register powerpoint x64
                        RegistryKey powerpnt14Export = officeRegistryBase.CreateSubKey(@"Office\14.0\PowerPoint\Presentation Converters\OOXML Converters\Export\UOF Presentation Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        powerpnt14Export.SetValue("Clsid", "{AEC7E6B1-3D6E-41E0-A442-34A0D6995DA4}", RegistryValueKind.String);
                        powerpnt14Export.SetValue("Extensions", "uop", RegistryValueKind.String);
                        powerpnt14Export.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        RegistryKey powerpnt14Import = officeRegistryBase.CreateSubKey(@"Office\14.0\PowerPoint\Presentation Converters\OOXML Converters\Import\UOF Presentation Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        powerpnt14Import.SetValue("Clsid", "{AEC7E6B1-3D6E-41E0-A442-34A0D6995DA4}", RegistryValueKind.String);
                        powerpnt14Import.SetValue("Extensions", "uop", RegistryValueKind.String);
                        powerpnt14Import.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        // register word x64
                        RegistryKey word14Export = officeRegistryBase.CreateSubKey(@"Office\14.0\Word\Text Converters\OOXML Converters\Export\UOF Word Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        word14Export.SetValue("Clsid", "{6970F0A4-88D6-4BAE-BC69-5C40E3598336}", RegistryValueKind.String);
                        word14Export.SetValue("Extensions", "uot", RegistryValueKind.String);
                        word14Export.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        RegistryKey word14Import = officeRegistryBase.CreateSubKey(@"Office\14.0\Word\Text Converters\OOXML Converters\Import\UOF Word Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        word14Import.SetValue("Clsid", "{6970F0A4-88D6-4BAE-BC69-5C40E3598336}", RegistryValueKind.String);
                        word14Import.SetValue("Extensions", "uot", RegistryValueKind.String);
                        word14Import.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        // register excel x64
                        RegistryKey excel14Export = officeRegistryBase.CreateSubKey(@"Office\14.0\Excel\Text Converters\OOXML Converters\Export\UOF Spreadsheet Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        excel14Export.SetValue("Clsid", "{ADBE850B-D37F-4422-B66E-88471BDC1B20}", RegistryValueKind.String);
                        excel14Export.SetValue("Extensions", "uos", RegistryValueKind.String);
                        excel14Export.SetValue("Name", "UOF2.0", RegistryValueKind.String);

                        RegistryKey excel14Import = officeRegistryBase.CreateSubKey(@"Office\14.0\Excel\Text Converters\OOXML ConvertersImport\UOF Spreadsheet Converter_50", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryOptions.None);
                        excel14Import.SetValue("Clsid", "{ADBE850B-D37F-4422-B66E-88471BDC1B20}", RegistryValueKind.String);
                        excel14Import.SetValue("Extensions", "uos", RegistryValueKind.String);
                        excel14Import.SetValue("Name", "UOF2.0", RegistryValueKind.String);
                    }
                }
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Success;
                //return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        [CustomAction]
        public static ActionResult X64OfficeUnRegister(Session session)
        {
            session.Log("Begin removing register UOFTranslator X64 Office");
            Record record = new Record(0);

            try
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    RegistryKey registryBase = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);

                    string regpath = @"SOFTWARE\Microsoft\";
                    RegistryKey officeRegistryBase = registryBase.OpenSubKey(regpath,
                        RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.FullControl);

                    if (officeRegistryBase != null)
                    {
                        // unregister powerpoint x64
                        officeRegistryBase.DeleteSubKeyTree(@"Office\15.0\PowerPoint\Presentation Converters\OOXML Converters\Export\UOF Presentation Converter_50", false);

                        officeRegistryBase.DeleteSubKeyTree(@"Office\15.0\PowerPoint\Presentation Converters\OOXML Converters\Import\UOF Presentation Converter_50", false);

                        // Unregister word x64
                        officeRegistryBase.DeleteSubKeyTree(@"Office\15.0\Word\Text Converters\OOXML Converters\Export\UOF Word Converter_50", false);

                        officeRegistryBase.DeleteSubKeyTree(@"Office\15.0\Word\Text Converters\OOXML Converters\Import\UOF Word Converter_50", false);

                        // Unregister excel x64
                        officeRegistryBase.DeleteSubKeyTree(@"Office\15.0\Excel\Text Converters\OOXML Converters\Export\UOF Spreadsheet Converter_50", false);

                        officeRegistryBase.DeleteSubKeyTree(@"Office\15.0\Excel\Text Converters\OOXML ConvertersImport\UOF Spreadsheet Converter_50", false);


                        // unsupport office 2010 x64
                        // unregister powerpoint x64
                        officeRegistryBase.DeleteSubKeyTree(@"Office\14.0\PowerPoint\Presentation Converters\OOXML Converters\Export\UOF Presentation Converter_50", false);

                        officeRegistryBase.DeleteSubKeyTree(@"Office\14.0\PowerPoint\Presentation Converters\OOXML Converters\Import\UOF Presentation Converter_50", false);

                        // Unregister word x64
                        officeRegistryBase.DeleteSubKeyTree(@"Office\14.0\Word\Text Converters\OOXML Converters\Export\UOF Word Converter_50", false);

                        officeRegistryBase.DeleteSubKeyTree(@"Office\14.0\Word\Text Converters\OOXML Converters\Import\UOF Word Converter_50", false);

                        // Unregister excel x64
                        officeRegistryBase.DeleteSubKeyTree(@"Office\14.0\Excel\Text Converters\OOXML Converters\Export\UOF Spreadsheet Converter_50", false);

                        officeRegistryBase.DeleteSubKeyTree(@"Office\14.0\Excel\Text Converters\OOXML ConvertersImport\UOF Spreadsheet Converter_50", false);
                    }
                }
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Success;
                //return ActionResult.Failure;
            }
            return ActionResult.Success;
        }
    }
}
