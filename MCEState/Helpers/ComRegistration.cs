/*
 * Copyright (c) 2007 Jonathan Bradshaw
 * 
 */
using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.Runtime.InteropServices;
using System.Collections;
using System.Diagnostics;

namespace VmcController.MceState.Helpers
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
            RegistrationServices regServices = new RegistrationServices();
            if (!regServices.RegisterAssembly(GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase))
                throw new InstallException("Error during registration of " + GetType().Assembly.FullName);
        }

        /// <summary>
        /// When overridden in a derived class, completes the install transaction.
        /// </summary>
        /// <param name="savedState">An <see cref="T:System.Collections.IDictionary"/> that contains the state of the computer after all the installers in the collection have run.</param>
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
            //  Kill any ehmsas processes so our sink will be used on relaunch
            foreach (Process proc in Process.GetProcessesByName("ehmsas"))
            {
                proc.Kill();
                proc.Dispose();
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
                RegistrationServices regServices = new RegistrationServices();
                regServices.UnregisterAssembly(GetType().Assembly);
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
                RegistrationServices regServices = new RegistrationServices();
                regServices.UnregisterAssembly(GetType().Assembly);
            }
            catch { }
            finally
            {
                base.Uninstall(savedState);
            }
        }
    }

}
