/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.Collections;
using System.EnterpriseServices;
using System.IO;

namespace VmcController.Services
{
    /// <summary>
    /// Custom Installer Class for Registering this Assembly for COM Interop
    /// </summary>
    [RunInstaller(true)]
    public partial class ComRegistration : Installer
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ComRegistration"/> class.
        /// </summary>
        public ComRegistration()
        {
            InitializeComponent();
        }

        /// <summary>
        /// When overridden in a derived class, performs the installation.
        /// </summary>
        /// <param name="stateSaver">An <see cref="T:System.Collections.IDictionary"/> used to save information needed to perform a commit, rollback, or uninstall operation.</param>
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            string appID = string.Empty;
            string typeLib = string.Empty;
            // Get the location of the current assembly
            string assemblyLocation = GetType().Assembly.Location;
            // Install the application
            try
            {
                RegistrationHelper regHelper = new RegistrationHelper();
                regHelper.InstallAssembly(assemblyLocation, ref appID, ref typeLib, InstallationFlags.FindOrCreateTargetApplication);
                // Save the state - you will need this for the uninstall
                stateSaver.Add("AppID", appID);
                stateSaver.Add("Assembly", assemblyLocation);
            }
            catch (Exception ex)
            {
                throw new InstallException("Error during registration of " + GetType().Assembly.FullName, ex);
            }
        }

        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
            try
            {
                INetFwMgr mgr = (INetFwMgr)new NetFwMgr();
                INetFwProfile profile = mgr.LocalPolicy.CurrentProfile;
                string winDir = System.Environment.GetFolderPath(Environment.SpecialFolder.System);
                winDir = winDir.Substring(0, winDir.LastIndexOf(Path.DirectorySeparatorChar));

                INetFwAuthorizedApplication fwApp = (INetFwAuthorizedApplication)new NetFwAuthorizedApplication();
                fwApp.Name = "Media Center Extensibility Host";
                fwApp.ProcessImageFileName = winDir + @"\ehome\ehexthost.exe";
                fwApp.Enabled = true;
                fwApp.IpVersion = IPVersion.IPAny;
                fwApp.Scope = Scope.Subnet;
                fwApp.RemoteAddresses = "*";

                profile.AuthorizedApplications.Add(fwApp);
            }
            catch (Exception ex)
            {
                throw new InstallException("Error during firewall registration of ehexthost.exe", ex);
            }

            try
            {
                INetFwMgr mgr = (INetFwMgr)new NetFwMgr();
                INetFwProfile profile = mgr.LocalPolicy.CurrentProfile;
                string winDir = System.Environment.GetFolderPath(Environment.SpecialFolder.System);
                winDir = winDir.Substring(0, winDir.LastIndexOf(Path.DirectorySeparatorChar));

                INetFwAuthorizedApplication fwApp = (INetFwAuthorizedApplication)new NetFwAuthorizedApplication();
                fwApp.Name = "Media Center Media Status Aggregator Service";
                fwApp.ProcessImageFileName = winDir + @"\ehome\ehmsas.exe";
                fwApp.Enabled = true;
                fwApp.IpVersion = IPVersion.IPAny;
                fwApp.Scope = Scope.Subnet;
                fwApp.RemoteAddresses = "*";

                profile.AuthorizedApplications.Add(fwApp);
            }
            catch (Exception ex)
            {
                throw new InstallException("Error during firewall registration of ehmsas.exe", ex);
            }
        }

        /// <summary>
        /// When overridden in a derived class, restores the pre-installation state of the computer.
        /// </summary>
        /// <param name="savedState">An <see cref="T:System.Collections.IDictionary"/> that contains the pre-installation state of the computer.</param>
        public override void Rollback(IDictionary savedState)
        {
            try
            {
                // Get the state created when the app was installed
                string appID = (string)savedState["AppID"];
                string assembly = (string)savedState["Assembly"];
                // Uninstall the application
                RegistrationHelper regHelper = new RegistrationHelper();
                regHelper.UninstallAssembly(assembly, appID);
            }
            catch (Exception) { }
            finally
            {
                base.Rollback(savedState);
            }
        }

        /// <summary>
        /// When overridden in a derived class, removes an installation.
        /// </summary>
        /// <param name="savedState">An <see cref="T:System.Collections.IDictionary"/> that contains the state of the computer after the installation was complete.</param>
        public override void Uninstall(IDictionary savedState)
        {
            try
            {
                // Get the state created when the app was installed
                string appID = (string)savedState["AppID"];
                string assembly = (string)savedState["Assembly"];
                // Uninstall the application
                RegistrationHelper regHelper = new RegistrationHelper();
                regHelper.UninstallAssembly(assembly, appID);
            }
            catch (Exception) { }
            finally
            {
                base.Uninstall(savedState);
            }
        }
    }

}
