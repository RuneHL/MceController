/*
 * This module is build on top of on J.Bradshaw's vmcController
 * Implements custom formatting of command output for the html server
 * 
 * Copyright (c) 2008 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * History:
 * 2009-03-18 Created by Anthony Jones
 * 
 */

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace VmcController.AddIn.Commands
{
    /// <summary>
    /// Summary description for getArtists commands.
    /// </summary>
    public class customCmd : ICommand
    {
        private string m_input = "";
        private Dictionary<string, string> m_templates = new Dictionary<string, string>();

        public customCmd()
        {
            loadTemplate();
        }

        public customCmd(string sInputPage)
        {
            m_input = sInputPage;
            loadTemplate();
        }


        private bool loadTemplate()
        {
            bool ret = true;

            try
            {
                Regex re = new Regex("(?<lable>^\\w+?)\t+(?<format_file>.*$?)");
                StreamReader fTemplate = File.OpenText("custom.template");
                string sIn = null;
                while ((sIn = fTemplate.ReadLine()) != null)
                {
                    Match match = re.Match(sIn);
                    if (match.Success) m_templates.Add(match.Groups["lable"].Value, match.Groups["format_file"].Value);
                }
                fTemplate.Close();
            }
            catch (Exception) { ret = false; }

            return ret;
        }

        private string listTemplates()
        {
            string sRet = "";
            foreach (KeyValuePair<string, string> template in m_templates)
            {
                sRet += "     " + template.Key + ": " + template.Value + "\r\n";
            }
            if (sRet.Length == 0) sRet = "     -- No Templates Defined! --";
            return sRet;
        }

        private string loadTemplateFile(string fn)
        {
            string sRet = "";

            try
            {
                StreamReader fTemplate = File.OpenText(fn);
                sRet = fTemplate.ReadToEnd();
                fTemplate.Close();
            }
            catch (Exception) { sRet = "Error: Could not load '" + fn + "'"; }

            return sRet;
        }

        private string getRex(string rex)
        {
            Regex r = new Regex(rex);
            Match match = r.Match(m_input);
            if (match.Success) return match.ToString();
            return "";
        }

        private static Regex index_regex = new Regex("<(?<index>\\d+)>");

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            string s = "";
            s = "[-help | <custom format name>] - HTML interface only - applies custom format to previously executed nested commands";
            return s;
        }

        public OpResult showHelp(OpResult or)
        {
            or.AppendFormat("Currently defined templates:\r\n" + listTemplates());
            or.AppendFormat("Custom Formatting Notes:");
            or.AppendFormat("     All custom formats must be defined in the \"custom.template\" file");
            or.AppendFormat("     The \"custom.template\" file must be manually coppied to the ehome directory (usually C:\\Windows\\ehome)");
            or.AppendFormat("     The \"custom.template\" file contains notes / examples on formatting");

            return or;
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            bool bTemplate = true;
            OpResult opResult = new OpResult();
            string page_file = "";
            string page = "";

            opResult.StatusCode = OpStatusCode.Ok;
            try
            {
                if (param.IndexOf("-help") >= 0)
                {
                    opResult = showHelp(opResult);
                    return opResult;
                }

                // Use custom format?
                if (!m_templates.ContainsKey(param))
                {
                    bTemplate = false;
                    page = "Error: Template '" + param + "' not found!\r\n All available templates:\r\n";
                    page += listTemplates();
                }
                else page_file = m_templates[param];
                if (page_file.Length > 0) page = loadTemplateFile(page_file);

                // Convert tags:
                if (bTemplate)
                {
                    Regex rTags = new Regex("%%(?<rex>.+?)%%");
                    Match match = rTags.Match(page);
                    while (match.Success)
                    {
                        string value = getRex(match.Groups["rex"].Value);
                        page = page.Replace("%%" + match.Groups["rex"].Value + "%%", value);
                        match = match.NextMatch();
                    }
                }

                opResult.AppendFormat("{0}", page);
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        #endregion
    }
}
