namespace CppTripleSlash
{
    using EnvDTE;
    using Microsoft.VisualStudio;
    using Microsoft.VisualStudio.Language.Intellisense;
    using Microsoft.VisualStudio.OLE.Interop;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Text;
    using Microsoft.VisualStudio.Text.Editor;
    using Microsoft.VisualStudio.TextManager.Interop;
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    public class TripleSlashCompletionCommandHandler : IOleCommandTarget
    {
        private const string CppTypeName = "C/C++";
        private IOleCommandTarget m_nextCommandHandler;
        private IWpfTextView m_textView;
        private TripleSlashCompletionHandlerProvider m_provider;
        private ICompletionSession m_session;
        DTE m_dte;

        public TripleSlashCompletionCommandHandler(
            IVsTextView textViewAdapter,
            IWpfTextView textView,
            TripleSlashCompletionHandlerProvider provider,
            DTE dte)
        {
            this.m_textView = textView;
            this.m_provider = provider;
            this.m_dte = dte;

            // add the command to the command chain
            if (textView != null &&
                textView.TextBuffer != null &&
                textView.TextBuffer.ContentType.TypeName == CppTypeName)
            {
                textViewAdapter.AddCommandFilter(this, out m_nextCommandHandler);
            }
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            return m_nextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (VsShellUtilities.IsInAutomationFunction(m_provider.ServiceProvider))
            {
                return m_nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }

            // make a copy of this so we can look at it after forwarding some commands 
            uint commandID = nCmdID;
            char typedChar = char.MinValue;

            // make sure the input is a char before getting it 
            if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
            {
                typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
            }

            // check for the triple slash
            if (typedChar == '/' && m_dte != null)
            {
                string currentLine = m_textView.TextSnapshot.GetLineFromPosition(
                    m_textView.Caret.Position.BufferPosition.Position).GetText();
                if ((currentLine + "/").Trim() == "///")
                {
                    // Calculate how many spaces
                    string spaces = currentLine.Replace(currentLine.TrimStart(), "");
                    TextSelection ts = m_dte.ActiveDocument.Selection as TextSelection;
                    int oldLine = ts.ActivePoint.Line;
                    int oldOffset = ts.ActivePoint.LineCharOffset;
                    ts.LineDown();
                    ts.EndOfLine();

                    CodeElement codeElement = null;
                    CodeFunction function = null;
                    CodeType codeType = null;
                    CodeVariable codeVariable = null;
                    CodeProperty codeProperty = null;

                    FileCodeModel fcm = m_dte.ActiveDocument.ProjectItem.FileCodeModel;
                    if (fcm != null)
                    {
                        codeElement = fcm.CodeElementFromPoint(ts.ActivePoint, vsCMElement.vsCMElementFunction);
                    }

                    if (codeElement != null)
                    {
                        function = codeElement as CodeFunction;
                        codeType = codeElement as CodeType;
                        codeVariable = codeElement as CodeVariable;
                        codeProperty = codeElement as CodeProperty;
                    }

                    if (function != null)
                    {
                        StringBuilder sb = new StringBuilder("/ <summary>\r\n" + spaces + "/// \r\n" + spaces + "/// </summary>");
                        foreach (CodeElement child in codeElement.Children)
                        {
                            CodeParameter parameter = child as CodeParameter;
                            if (parameter != null)
                            {
                                sb.AppendFormat("\r\n" + spaces + "/// <param name=\"{0}\"></param>", parameter.Name);
                            }
                        }

                        if (function.Type.AsString != "void")
                        {
                            sb.AppendFormat("\r\n" + spaces + "/// <returns></returns>");
                        }

                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.Insert(sb.ToString());
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.LineDown();
                        ts.EndOfLine();
                        return VSConstants.S_OK;
                    }
                    else if (codeType != null)
                    {
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.Insert("/ <summary>\r\n" + spaces + "/// \r\n" + spaces + "/// </summary>");
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.LineDown();
                        ts.EndOfLine();
                        return VSConstants.S_OK;
                    }
                    else if (codeVariable != null)
                    {
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.Insert("/ <summary>\r\n" + spaces + "/// \r\n" + spaces + "/// </summary>");
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.LineDown();
                        ts.EndOfLine();
                        return VSConstants.S_OK;
                    }
                    else if (codeProperty != null)
                    {
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.Insert("/ <summary>\r\n" + spaces + "/// \r\n" + spaces + "/// </summary>");
                        ts.MoveToLineAndOffset(oldLine, oldOffset);
                        ts.LineDown();
                        ts.EndOfLine();
                        return VSConstants.S_OK;
                    }
                    else if (codeElement.Kind == vsCMElement.vsCMElementVCBase)
                    {
                        CodeType parent = codeElement.Collection.Parent as CodeType;
                        if (parent != null)
                        {
                            ts.MoveToLineAndOffset(oldLine, oldOffset);
                            ts.Insert("/ <summary>\r\n" + spaces + "/// \r\n" + spaces + "/// </summary>");
                            ts.MoveToLineAndOffset(oldLine, oldOffset);
                            ts.LineDown();
                            ts.EndOfLine();
                            return VSConstants.S_OK;
                        }
                    }

                    ts.MoveToLineAndOffset(oldLine, oldOffset);
                }
            }

            if (m_session != null && !m_session.IsDismissed)
            {
                // check for a commit character 
                if (nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN
                    || nCmdID == (uint)VSConstants.VSStd2KCmdID.TAB
                    || typedChar == '>')
                {
                    // check for a selection 
                    // if the selection is fully selected, commit the current session 
                    if (m_session.SelectedCompletionSet.SelectionStatus.IsSelected)
                    {
                        m_session.Commit();
                        TextSelection ts = m_dte.ActiveDocument.Selection as TextSelection;
                        string currentLine = m_textView
                            .Caret
                            .ContainingTextViewLine
                            .Extent
                            .GetText();
                        if (currentLine.EndsWith("</param>"))
                        {
                            ts.CharLeft(false, 10);
                        }
                        else if (currentLine.EndsWith("-->"))
                        {
                            ts.CharLeft(false, 3);
                        }
                        else if (currentLine.EndsWith("]]>"))
                        {
                            ts.CharLeft(false, 3);
                        }
                        else if (currentLine.EndsWith("</example>"))
                        {
                            ts.CharLeft(false, 10);
                        }
                        else if (currentLine.EndsWith("</exception>"))
                        {
                            ts.CharLeft(false, 14);
                        }
                        else if (currentLine.EndsWith("</permission>"))
                        {
                            ts.CharLeft(false, 15);
                        }
                        else if (currentLine.EndsWith("</remarks>"))
                        {
                            ts.CharLeft(false, 10);
                        }
                        else if (currentLine.EndsWith("</returns>"))
                        {
                            ts.CharLeft(false, 10);
                        }
                        else if (currentLine.EndsWith("<see cref=\"\"/>"))
                        {
                            ts.CharLeft(false, 3);
                        }
                        else if (currentLine.EndsWith("<seealso cref=\"\"/>"))
                        {
                            ts.CharLeft(false, 3);
                        }
                        else if (currentLine.EndsWith("</value>"))
                        {
                            ts.CharLeft(false, 8);
                        }
                        else if (currentLine.EndsWith("</list>"))
                        {
                            ts.CharLeft(false, 7);
                        }
                        else if (currentLine.EndsWith("</c>"))
                        {
                            ts.CharLeft(false, 4);
                        }
                        else if (currentLine.EndsWith("</code>"))
                        {
                            ts.CharLeft(false, 7);
                        }
                        else if (currentLine.EndsWith("[@name=\"\"]'/>"))
                        {
                            ts.CharLeft(false, 21);
                        }

                        // also, don't add the character to the buffer 
                        return VSConstants.S_OK;
                    }
                    else
                    {
                        // if there is no selection, dismiss the session
                        m_session.Dismiss();
                    }
                }
            }
            else
            {
                if (pguidCmdGroup == VSConstants.VSStd2K && nCmdID == (uint)VSConstants.VSStd2KCmdID.RETURN)
                {
                    string currentLine = m_textView.TextSnapshot.GetLineFromPosition(
                            m_textView.Caret.Position.BufferPosition.Position).GetText();
                    if (currentLine.TrimStart().StartsWith("///"))
                    {
                        TextSelection ts = m_dte.ActiveDocument.Selection as TextSelection;
                        ts.NewLine();
                        ts.Insert("/// ");
                        return VSConstants.S_OK;
                    }
                }
            }

            // pass along the command so the char is added to the buffer
            int retVal = m_nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            if (typedChar == '<')
            {
                string currentLine = m_textView.TextSnapshot.GetLineFromPosition(
                            m_textView.Caret.Position.BufferPosition.Position).GetText();
                if (currentLine.TrimStart().StartsWith("///"))
                {
                    if (m_session == null || m_session.IsDismissed) // If there is no active session, bring up completion
                    {
                        if (this.TriggerCompletion())
                        {
                            m_session.Filter();
                            return VSConstants.S_OK;
                        }
                    }
                }
            }
            else if (
                commandID == (uint)VSConstants.VSStd2KCmdID.BACKSPACE ||
                commandID == (uint)VSConstants.VSStd2KCmdID.DELETE ||
                char.IsLetter(typedChar))
            {
                if (m_session != null && !m_session.IsDismissed) // the completion session is already active, so just filter
                {
                    m_session.Filter();
                    return VSConstants.S_OK;
                }
            }

            return retVal;
        }

        private bool TriggerCompletion()
        {
            if (m_session != null)
            {
                return false;
            }

            // the caret must be in a non-projection location 
            SnapshotPoint? caretPoint =
            m_textView.Caret.Position.Point.GetPoint(
                textBuffer => (!textBuffer.ContentType.IsOfType("projection")), PositionAffinity.Predecessor);
            if (!caretPoint.HasValue)
            {
                return false;
            }

            m_session = m_provider.CompletionBroker.CreateCompletionSession(
                m_textView,
                caretPoint.Value.Snapshot.CreateTrackingPoint(caretPoint.Value.Position, PointTrackingMode.Positive),
                true);

            // subscribe to the Dismissed event on the session 
            m_session.Dismissed += this.OnSessionDismissed;
            m_session.Start();
            return true;
        }

        private void OnSessionDismissed(object sender, EventArgs e)
        {
            m_session.Dismissed -= this.OnSessionDismissed;
            m_session = null;
        }
    }
}