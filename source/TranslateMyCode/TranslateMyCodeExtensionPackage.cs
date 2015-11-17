using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Editor;
using TranslateMyCode.Translate;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TomiNiemenmaa.TranslateMyCode
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidTranslateMyCodePkgString)]
    // This attribute is needed to let the shell know that this package exposes options page.
    [ProvideOptionPage(typeof(OptionsPageGrid), "TranslateMyCode", "Bing Translation", 0, 0, true)]
    public sealed class TranslateMyCodePackage : Package
    {
        // BingTranslator could be replaced by another by another class
        // that impelements ITranslator interface
        private ITranslator _translator = null;
        
        // Getter for source language from options
        private Language SourceLanguage
        {
            get
            {
                var optionsPage = (OptionsPageGrid)GetDialogPage(typeof(OptionsPageGrid));
                return optionsPage.SourceLanguage;
            }
        }

        // Getter for target language from options
        private Language TargetLanguage
        {
            get
            {
                var optionsPage = (OptionsPageGrid)GetDialogPage(typeof(OptionsPageGrid));
                return optionsPage.TargetLanguage;
            }
        }

        // Initialize package and attach menu item to Translate command
        protected override void Initialize()
        {
            base.Initialize();

            // Add command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;

            if (null != mcs)
            {
                // Register event handler for Translate command on right click menu
                CommandID menuCommandID = new CommandID(GuidList.guidTranslateMyCodeCmdSet, (int)PkgCmdIDList.cmdTranslate);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                mcs.AddCommand(menuItem);
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            // Let exceptions flow here so we can log them to ActivityLog
            try
            {
                if (_translator == null)
                {
                    var settings = (OptionsPageGrid)GetDialogPage(typeof(OptionsPageGrid));
                    _translator = new BingTranslator(settings.BingClientID, settings.BingClientSecret);
                }
                // Find the active windows host
                IWpfTextViewHost host = GetCurrentViewHost();

                // Get highlighted selection from window
                ITextSelection selection = host.TextView.Selection;

                if (selection == null) return;

                // Get snapshot of the selection
                ITextSnapshot snapshot = selection.StreamSelectionSpan.Snapshot;

                if (snapshot == null) return;

                ITextBuffer buffer = snapshot.TextBuffer;

                using (ITextEdit edit = buffer.CreateEdit())
                {

                    if (edit != null)
                    {
                        // Get the highlighted text
                        string wordToTranslate = selection.StreamSelectionSpan.GetText();

                        string word = _translator.Translate(wordToTranslate, SourceLanguage, TargetLanguage);
                        word = FormatTranslation(word);

                        // Get starting position and length for selected text so we can replace
                        // the text in the active window with the translation
                        int start = selection.StreamSelectionSpan.Start.Position.Position;
                        int length = selection.StreamSelectionSpan.Length;

                        // Replace the word in visual studio
                        edit.Delete(start, length);
                        edit.Insert(start, word);
                        edit.Apply();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was error trying to get translation: " + ex.Message);
                Microsoft.VisualStudio.Shell.ActivityLog.LogError(ex.Source, ex.Message);
            }
        }

        // Words that will be removed from translation
        private readonly List<string> RedundantWords = new List<string>
        {
            "the",
            "a",
            "an"
        };
            
        /// Remove articles and if original word is using pascal casing use also else use camel casing
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string FormatTranslation(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return "";
            }
            
            input = input.Trim();
            
            string retVal = "";
            string[] words = input.Split(' ');
            bool firstWord = true;

            // If the first letter is uppercase use pascal casing for translation
            // else use camel casing
            bool PascalCasing = char.IsUpper(input[0]);
            foreach (string word in words)
            {
                if (!RedundantWords.Contains(word.ToLower()))
                {
                    string sWord = word.Trim();
                    if (firstWord)
                    {
                        retVal += PascalCasing ? char.ToUpper(sWord[0]) + (sWord.Length > 1 ? sWord.Substring(1) : "") : char.ToLower(sWord[0]) + (sWord.Length > 1 ? sWord.Substring(1) : "");
                    }
                    else
                    {
                        retVal += char.ToUpper(sWord[0]) + (sWord.Length > 1 ? sWord.Substring(1) : "");
                    }
                    firstWord = false;
                }
            }
            return retVal;
        }

        private IWpfTextViewHost GetCurrentViewHost()
        {
            IVsTextView vTextView;
            var textManager = (IVsTextManager)this.GetService(typeof(SVsTextManager));

            int mustHaveFocus = 1;
            textManager.GetActiveView(mustHaveFocus, null, out vTextView);

            IVsUserData userData = vTextView as IVsUserData;

            if (userData != null)
            {
                IWpfTextViewHost viewHost;
                object holder;
                Guid guidViewHost = DefGuidList.guidIWpfTextViewHost;
                userData.GetData(ref guidViewHost, out holder);
                viewHost = (IWpfTextViewHost)holder;
                return viewHost;
            }
            else
            {
                return null;
            }
        }
    }

    public class OptionsPageGrid : DialogPage
    {
        private Language _sourceLanguage = Language.FI;
        private Language _targetLanguage = Language.EN;
        private string _bingClientID = null;
        private string _bingClientSecret = null;

        [Category("Languages")]
        [DisplayName("Source Language")]
        [Description("Source language for translation")]
        public Language SourceLanguage
        {
            get { return _sourceLanguage; }
            set { _sourceLanguage = value; }
        }

        [Category("Languages")]
        [DisplayName("Target Language")]
        [Description("Target language for translation")]
        public Language TargetLanguage
        {
            get { return _targetLanguage; }
            set { _targetLanguage = value; }
        }

        [Category("Bing Translator API")]
        [DisplayName("Client ID")]
        [Description("Bing Translator API ClientID")]
        public string BingClientID
        {
            get { return _bingClientID; }
            set { _bingClientID = value; }
        }

        [Category("Bing Translator API")]
        [DisplayName("Client Secret")]
        [Description("Bing Translator API client secret")]
        [PasswordPropertyText(true)]
        public string BingClientSecret
        {
            get { return _bingClientSecret; }
            set { _bingClientSecret = value; }
        }


    }
}
