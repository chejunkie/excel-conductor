using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using xlApp = Microsoft.Office.Interop.Excel.Application;

namespace ChEJunkie.Office.Excel 
{
    /// <summary>
    /// Represents the collection of all Excel instances running in a specific Windows session.
    /// </summary>
    public class Session
    {
        /// <summary>
        /// Gets an instance representing the current Windows session.
        /// </summary>
        public static Session Current => new Session(Process.GetCurrentProcess().SessionId);

        /// <summary>
        /// Initializes a new instance of the <see cref="Session"/> class.
        /// </summary>
        /// <param name="sessionId">The session identifier.</param>
        public Session(int sessionId)
        {
            Debug.WriteLine("");
            Debug.WriteLine("Session.Constructor");

            SessionId = sessionId;
        }

        /// <summary>Gets a sequence of all processes in this session named "EXCEL".</summary>
        private IEnumerable<Process> Processes =>
            Process.GetProcessesByName("EXCEL")
            .Where(p => p.SessionId == this.SessionId);

        /// <summary>Tries to convert the given process to an Excel instance,
        /// but returns null if an exception is thrown.</summary>
        private static xlApp? TryGetApp(Process process)
        {
            try
            {
                return process.AsExcelApp();
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        /// <summary>
        /// Gets the session identifier.
        /// </summary>
        public int SessionId { get; }

        /// <summary> Gets a sequence of process IDs for all currently running Excel
        /// processes in the specified Windows session.</summary>
        public IEnumerable<int> ProcessIds =>
            Processes
            .Select(p => p.Id)
            .ToArray();

        /// <summary>Gets a sequence of process IDs for all currently running processes
        /// in the specified Windows session named Excel, but which can currently be
        /// converted to Application instances.</summary>
        public IEnumerable<int> ReachableProcessIds =>
            AppsImpl.Select(a => a.AsProcess().Id).ToArray();

        /// <summary>Gets a sequence of process IDs for all currently running processes
        /// in the specified Windows session named Excel, but which cannot currently be
        /// converted to Application instances.</summary>
        public IEnumerable<int> UnreachableProcessIds =>
            ProcessIds
            .Except(ReachableProcessIds)
            .ToArray();

        /// <summary>Gets a sequence of all currently accessible Excel instances running
        /// in the specified Windows session.</summary>
        /// <remarks>Sequence is untyped because Application is an embedded interop type.
        /// Elements can be safely cast to Application.</remarks>
        public IEnumerable Applications => AppsImpl; // https://learn.microsoft.com/en-us/dotnet/framework/interop/type-equivalence-and-embedded-interop-types

        /// <summary>Gets a strongly-typed sequence of all currently accessible Excel instances
        /// running in the specified Windows session.</summary>
#pragma warning disable CS8619 // Nullability of reference types in value doesn't match target type.
        private IEnumerable<xlApp> AppsImpl =>
            Processes
            .Select(TryGetApp)
            .Where(a => a != null && a.AsProcess().IsVisible())
            .ToArray();
#pragma warning restore CS8619 // Nullability of reference types in value doesn't match target type.

        /// <summary>Gets the Excel instance with the topmost window.
        /// Returns null if no accessible instances.</summary>
        public xlApp? TopMost
        {
            get
            {
                var dict = AppsImpl.ToDictionary(
                    keySelector: a => a.AsProcess(),
                    elementSelector: a => a);

                var topProcess = dict.Keys.TopMost();

                if (topProcess == null)
                {
                    return null;
                }
                else
                {
                    try
                    {
                        return dict[topProcess];
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
        }

        /// <summary>Gets the "default" Excel instance that double-clicked
        /// files will open in. Returns null if no accesible instances.</summary>
        public xlApp? PrimaryInstance
        {
            get
            {
                try
                {
                    return (xlApp)WinApi.GetActiveObject("Excel.Application");
                }
                catch (COMException x)
                when (x.Message.StartsWith("Operation unavailable"))
                {
                    Debug.WriteLine("Session: Primary Excel instance unavailable.");
                    return null;
                }
            }
        }
    }
}