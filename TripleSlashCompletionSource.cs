﻿namespace CppTripleSlash
{
    using Microsoft.VisualStudio.Language.Intellisense;
    using Microsoft.VisualStudio.Text;
    using Microsoft.VisualStudio.Text.Operations;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Media;

    public class TripleSlashCompletionSource : ICompletionSource
    {
        private TripleSlashCompletionSourceProvider m_sourceProvider;
        private ITextBuffer m_textBuffer;
        private List<Completion> m_compList = new List<Completion>();
        private bool m_isDisposed;
        
        public TripleSlashCompletionSource(TripleSlashCompletionSourceProvider sourceProvider, ITextBuffer textBuffer)
        {
            m_sourceProvider = sourceProvider;
            m_textBuffer = textBuffer;
            ImageSource image = this.m_sourceProvider.GlyphService.GetGlyph(StandardGlyphGroup.GlyphKeyword, StandardGlyphItem.GlyphItemPublic);
            m_compList.Add(new Completion("<param>", "<param name=\"\"></param>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<!-->", "<!---->", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<![CDATA[>", "<![CDATA[]]>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<example>", "<example></example>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<exception>", "<exception cref=\"\"></exception>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<include>", "<include file='' path='[@name=\"\"]'/>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<permission>", "<permission cref=\"\"></permission>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<remarks>", "<remarks></remarks>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<returns>", "<returns></returns>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<see>", "<see cref=\"\"/>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<seealso>", "<seealso cref=\"\"/>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<value>", "<value></value>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<list>", "<list></list>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<c>", "<c></c>", string.Empty, image, string.Empty));
            m_compList.Add(new Completion("<code>", "<code></code>", string.Empty, image, string.Empty));
        }

        void ICompletionSource.AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            if (m_isDisposed)
            {
                return;
            }

            SnapshotPoint? snapshotPoint = session.GetTriggerPoint(m_textBuffer.CurrentSnapshot);
            if (!snapshotPoint.HasValue)
            {
                return;
            }

            string text = snapshotPoint.Value.GetContainingLine().GetText();
            if (!text.TrimStart().StartsWith("///"))
            {
                return;
            }

            ITrackingSpan trackingSpan = FindTokenSpanAtPosition(session.GetTriggerPoint(m_textBuffer), session);
            var newCompletionSet = new CompletionSet(
                "TripleSlashCompletionSet",
                "TripleSlashCompletionSet",
                trackingSpan,
                m_compList,
                Enumerable.Empty<Completion>());
            completionSets.Add(newCompletionSet);
        }

        public void Dispose()
        {
            if (!m_isDisposed)
            {
                GC.SuppressFinalize(this);
                m_isDisposed = true;
            }
        }

        private ITrackingSpan FindTokenSpanAtPosition(ITrackingPoint point, ICompletionSession session)
        {
            SnapshotPoint currentPoint = session.TextView.Caret.Position.BufferPosition - 1;
            ITextStructureNavigator navigator = m_sourceProvider.NavigatorService.GetTextStructureNavigator(m_textBuffer);
            TextExtent extent = navigator.GetExtentOfWord(currentPoint);
            return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
        }
    }
}