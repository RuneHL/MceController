/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely, subject to the following restrictions:
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 */
using System;
using System.Text;
using Microsoft.MediaCenter;
using Microsoft.MediaCenter.Hosting;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for NavigateTo command.
	/// </summary>
	public class NavigateToPage : ICommand
	{
        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            StringBuilder sb = new StringBuilder("<");
            foreach (string value in Enum.GetNames(typeof(PageId)))
                sb.AppendFormat("{0}|", value);
            sb.Remove(sb.Length-1, 1);
            sb.Append("> <optional page parameters>");
            return sb.ToString();
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult();
            try
            {
                string[] args = param.Split(' ');
                PageId pageId = (PageId)Enum.Parse(typeof(PageId), args[0], true);
                object obj = (args.Length == 2 ? args[1] : null);
                AddInHost.Current.MediaCenterEnvironment.NavigateToPage(pageId, obj);
                opResult.StatusCode = OpStatusCode.Success;
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
