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
using System.Collections.Generic;
using Microsoft.MediaCenter.TV.Scheduling;
using System.Diagnostics;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for FullScreen command.
	/// </summary>
	public class ScheduleCmd: ICommand
	{
        private EventSchedule m_eventSchedule;
        private string m_exception;

        public ScheduleCmd()
        {
            try
            {
                m_eventSchedule = new EventSchedule();
            }
            catch (EventScheduleException ex)
            {
                m_exception = ex.Message;
                Trace.TraceError(ex.ToString());
            }
        }

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "<recording|recorded|scheduled>";
        }

        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            List<ScheduleEvent> events;
            StringBuilder sb = new StringBuilder();
            OpResult opResult = new OpResult();

            if (m_eventSchedule == null)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = m_exception;
                return opResult;
            }

            try
            {
                if (param.Equals("recorded", StringComparison.InvariantCultureIgnoreCase))
                    events = m_eventSchedule.GetScheduleEvents(DateTime.MinValue, DateTime.MaxValue, ScheduleEventStates.HasOccurred) as List<ScheduleEvent>;
                else if (param.Equals("recording", StringComparison.InvariantCultureIgnoreCase))
                    events = m_eventSchedule.GetScheduleEvents(DateTime.MinValue, DateTime.MaxValue, ScheduleEventStates.IsOccurring) as List<ScheduleEvent>;
                else if (param.Equals("scheduled", StringComparison.InvariantCultureIgnoreCase))
                    events = m_eventSchedule.GetScheduleEvents(DateTime.Now, DateTime.Now.AddDays(7), ScheduleEventStates.WillOccur) as List<ScheduleEvent>;
                else
                {
                    opResult.StatusCode = OpStatusCode.BadRequest;
                    return opResult;
                }

                events.Sort(CompareScheduleEvents);
                foreach (ScheduleEvent item in events)
                {
                    opResult.AppendFormat("{0}={1} ({2}-{3})",
                        item.Id, item.GetExtendedProperty("Title"),
                        item.StartTime.ToLocalTime().ToString("g"),
                        item.EndTime.ToLocalTime().ToShortTimeString()
                        );
                }
                opResult.StatusCode = OpStatusCode.Ok;
            }
            catch (Exception ex)
            {
                opResult.StatusCode = OpStatusCode.Exception;
                opResult.StatusText = ex.Message;
            }
            return opResult;
        }

        #endregion

        /// <summary>
        /// Compares the schedule events for a reverse sort
        /// </summary>
        /// <param name="x">The first event.</param>
        /// <param name="y">The second event.</param>
        /// <returns></returns>
        private static int CompareScheduleEvents(ScheduleEvent x, ScheduleEvent y)
        {
            if (x.StartTime > y.StartTime) return -1;
            if (x.StartTime < y.StartTime) return 1;
            return 0;
        }
    }
}
