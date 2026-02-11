using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace MethodologyPOC
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<string, string> _emailByDoc =
                   new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private readonly HashSet<string> _policyAppliedForDoc =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private string GetDocKey(Word.Document doc)
        {

            try { return string.IsNullOrWhiteSpace(doc.FullName) ? doc.Name : doc.FullName; }
            catch { return doc.Name; }
        }

        public static string CurrentEmail = "";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                
                Application.DocumentOpen += Application_DocumentOpen;
                Application.WindowActivate += Application_WindowActivate;
                Application.DocumentBeforeClose += Application_DocumentBeforeClose;

            }
            catch (Exception ex)
            {
                System.Console.WriteLine("ThisAddIn_Startup Error" + ex.Message);
            }
        }


        private string GetUserEmailForPoc()
        {
            return PromptForEmail();
        }


        private string PromptForEmail()
        {
            using (var form = new Form())
            using (var tb = new TextBox())
            using (var ok = new Button())
            using (var cancel = new Button())
            {
                form.Text = "POC Identity : Please Enter your Email Id ";
                form.Width = 420;
                form.Height = 140;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                tb.Left = 12; tb.Top = 12; tb.Width = 380;

                ok.Text = "OK";
                ok.Left = 230; ok.Top = 45; ok.Width = 75;
                ok.DialogResult = DialogResult.OK;

                cancel.Text = "Cancel";
                cancel.Left = 317; cancel.Top = 45; cancel.Width = 75;
                cancel.DialogResult = DialogResult.Cancel;

                form.Controls.Add(tb);
                form.Controls.Add(ok);
                form.Controls.Add(cancel);
                form.AcceptButton = ok;
                form.CancelButton = cancel;

                return form.ShowDialog() == DialogResult.OK ? tb.Text.Trim() : "";
            }
        }

        private async void Application_DocumentOpen(Word.Document doc)
        {

            if (!IsTargetDoc(doc)) return;
           

            try
            {
                // If protected, we must unprotect to insert at end
                bool wasProtected = doc.ProtectionType != Word.WdProtectionType.wdNoProtection;
                if (wasProtected) ForceUnprotect(doc);

                await WebHtmlEmbedder.InsertWebHtmlAtEndAsync(doc, ConfigHelper.WebDemoUrl, 60000);

                ApplyAccessPolicy(doc);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Application_DocumentOpen" + ex.Message);
            }

        }

        private void Application_DocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            if (doc == null) return;

            var key = GetDocKey(doc);
            _policyAppliedForDoc.Remove(key);
            _emailByDoc.Remove(key);
        }
        private bool IsTargetDoc(Word.Document doc)
        {
            if (doc == null) return false;

            // check file name 
            return doc.Name.Equals(
                ConfigHelper.TargetDocumentName,
                StringComparison.OrdinalIgnoreCase);
        }

        private void ApplyAccessPolicy(Word.Document doc)
        {
            if (!IsTargetDoc(doc)) return;

            var key = GetDocKey(doc);

            if (_policyAppliedForDoc.Contains(key))
                return;

            string email;
            if (!_emailByDoc.TryGetValue(key, out email) || string.IsNullOrWhiteSpace(email))
            {
                email = PromptForEmail();
                if (string.IsNullOrWhiteSpace(email)) return;
                _emailByDoc[key] = email;
            }

            // set for ribbon logic
            CurrentEmail = email;
            try { Globals.Ribbons.MethodologyRibbon.ApplyVisibility(); } 
            catch { System.Console.WriteLine("ApplyAccessPolicy"); }

            ClearAllEditableRanges(doc); // instead of ForceUnprotect + ClearAllEditors_AllStories

            if (email.Equals(ConfigHelper.User1Email, StringComparison.OrdinalIgnoreCase))
                ApplySingleBookmarkPolicy(doc, "BM_VersioningPractice");
            else if (email.Equals(ConfigHelper.User2Email, StringComparison.OrdinalIgnoreCase))
                ApplySingleBookmarkPolicy(doc, "BM_LanguageSupport");
            else
                MakeFullyEditable(doc);


            _policyAppliedForDoc.Add(key);

            int countAll = doc.Range().Editors.Count;
            int countV = doc.Bookmarks["BM_VersioningPractice"].Range.Editors.Count;
            int countL = doc.Bookmarks["BM_LanguageSupport"].Range.Editors.Count;


        }

       
        private void MakeFullyEditable(Word.Document doc)
        {
            ClearAllEditableRanges(doc);
            // doc stays unprotected and fully editable
        }


        private void ApplySingleBookmarkPolicy(Word.Document doc, string bookmarkName)
        {
            if (!doc.Bookmarks.Exists(bookmarkName))
            {
                MessageBox.Show($"Bookmark not found: {bookmarkName}");
                return;
            }

            ClearAllEditableRanges(doc);

            // Protect 
            doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, NoReset: false, Password: ConfigHelper.ProtectPassword);

            // Clear  AFTER protect 
            TryDeleteEditable(doc, Word.WdEditorType.wdEditorEveryone);
            TryDeleteEditable(doc, Word.WdEditorType.wdEditorCurrent);

            // Add  editable region
            Word.Range rng = null;
            try
            {
                rng = doc.Bookmarks[bookmarkName].Range;
                rng.Editors.Add(Word.WdEditorType.wdEditorCurrent);

                // keep bookmark 
                doc.Bookmarks.Add(bookmarkName, rng);
            }
            finally
            {
                if (rng != null) Marshal.FinalReleaseComObject(rng);
            }
        }


        private void ClearEditorsWholeDoc(Word.Document doc)
        {
            Word.Range whole = null;
            Word.Editors editors = null;

            try
            {
                whole = doc.Range(0, doc.Content.End);
                editors = whole.Editors;

                for (int i = editors.Count; i >= 1; i--)
                    editors.Item(i).Delete();
            }
            catch { System.Console.WriteLine("ClearEditorsWholeDoc"); }
            finally
            {
                if (editors != null) Marshal.FinalReleaseComObject(editors);
                if (whole != null) Marshal.FinalReleaseComObject(whole);
            }
        }


        private void Application_WindowActivate(Word.Document doc, Word.Window wn)
        {
            if (doc == null) return;
            if (!IsTargetDoc(doc)) return;
        }

        private void ProtectAndAllowBookmark(Word.Document doc, string bookmarkName)
        {
            if (!doc.Bookmarks.Exists(bookmarkName))
            {
                MessageBox.Show($"Bookmark not found: {bookmarkName}");
                return;
            }

            // Reset protection 
            doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, NoReset: false, Password: ConfigHelper.ProtectPassword);

            Word.Range rng = null;
            try
            {
                // Fetch after protection 
                rng = doc.Bookmarks[bookmarkName].Range;

                // Allow edits only in this range
                rng.Editors.Add(Word.WdEditorType.wdEditorCurrent);

            }
            finally
            {
                if (rng != null) Marshal.FinalReleaseComObject(rng);
            }
        }

        private void ForceUnprotect(Word.Document doc)
        {
            if (doc == null) return;

            if (doc.ProtectionType == Word.WdProtectionType.wdNoProtection)
                return;

            try
            {
                doc.Unprotect(ConfigHelper.ProtectPassword);
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("ForceUnprotect" + ex.Message);
            }
        }


        private void ClearAllEditors(Word.Document doc)
        {
            try
            {
                foreach (Word.Range story in doc.StoryRanges)
                {
                    Word.Range r = story;
                    while (r != null)
                    {
                        Word.Editors editors = null;
                        try
                        {
                            editors = r.Editors;
                            for (int i = editors.Count; i >= 1; i--)
                                editors.Item(i).Delete();
                        }
                        catch { System.Console.WriteLine("ClearAllEditors"); }
                        finally
                        {
                            if (editors != null) Marshal.FinalReleaseComObject(editors);
                        }

                        // Next linked story (headers/footers chain)
                        var next = r.NextStoryRange;
                        Marshal.FinalReleaseComObject(r);
                        r = next;
                    }
                }
            }
            catch { System.Console.WriteLine("ClearAllEditors"); }
        }

        private void ClearAllEditableRanges(Word.Document doc)
        {
            if (doc == null) return;

            // Must be unprotected 
            ForceUnprotect(doc);

            // Delete for common editor types
            TryDeleteEditable(doc, Word.WdEditorType.wdEditorEveryone);
            TryDeleteEditable(doc, Word.WdEditorType.wdEditorCurrent);
            TryDeleteEditable(doc, Word.WdEditorType.wdEditorOwners);
            TryDeleteEditable(doc, Word.WdEditorType.wdEditorEditors);

        }

        private void TryDeleteEditable(Word.Document doc, Word.WdEditorType t)
        {
            try { doc.DeleteAllEditableRanges(t); }
            catch { System.Console.WriteLine("TryDeleteEditable"); }
        }



        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
