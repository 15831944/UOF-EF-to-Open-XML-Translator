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
using System.Configuration.Install;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace UofCmdConverter
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      提供COM安装与卸载操作
    ///     Author:         邓追
    ///     Create Date:    2007-09-05
    /// </summary>
    [RunInstaller(true)]
    public class ComInstaller : Installer
    {
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);

            // Modify PATH environment variable
            try
            {
                string path = System.Environment.GetEnvironmentVariable("Path", System.EnvironmentVariableTarget.Machine);
                path += ";";
                path += System.IO.Path.GetDirectoryName(this.GetType().Assembly.Location);
                System.Environment.SetEnvironmentVariable("Path", path, System.EnvironmentVariableTarget.Machine);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            base.Uninstall(savedState);

            // Modify PATH environment variable
            try
            {
                string path = System.Environment.GetEnvironmentVariable("Path", System.EnvironmentVariableTarget.Machine);
                string targetdir = ";" + System.IO.Path.GetDirectoryName(this.GetType().Assembly.Location);
                path = path.Replace(targetdir, "");
                System.Environment.SetEnvironmentVariable("Path", path, System.EnvironmentVariableTarget.Machine);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

    }
}
