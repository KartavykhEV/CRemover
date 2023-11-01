using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.UI.Design;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Debugger;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace CommentRemover
{
    internal sealed class RemoveAllCommentsInSolCommand : BaseCommand<RemoveAllCommentsInSolCommand>
    {
        protected override void SetupCommands()
        {
            RegisterCommand(PackageGuids.guidPackageCmdSet, PackageIds.RemoveAllCommentsInSol);
        }


        internal void clearCurFile()
        {
            var view = ProjectHelpers.GetCurentTextView();
            var mappingSpans = GetClassificationSpans(view, "comment");

            if (!mappingSpans.Any())
                return;

            try
            {
                DeleteFromBuffer(view, mappingSpans);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex);
            }
            //finally
            //{
            //    DTE.UndoContext.Close();
            //}

        }

        /// <summary>
        /// Обработка проекта или директории проекта
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        private List<ProjectItem> getFiles(ProjectItems items)
        {
            if (items == null) return new List<ProjectItem>();
            List<ProjectItem> fls = new List<ProjectItem>();
            foreach (ProjectItem item in items)
            {
                if (item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile)
                    fls.Add(item);
                else
                    fls.AddRange(getFiles(item.ProjectItems));
            }
            return fls;
        }
        /// <summary>
        /// Обработка директории решения
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private List<ProjectItem> processSolutionFolder(Project p)
        {
            List<ProjectItem> fls = new List<ProjectItem>();
            foreach (ProjectItem pi in p.ProjectItems)
            {
                Project pp = pi.Object as Project;
                if (pp == null) continue;

                fls.AddRange(pp.Kind == EnvDTE.Constants.vsProjectKindSolutionItems ? processSolutionFolder(pp) : getFiles(pp.ProjectItems));
            }
            return fls;
        }
        protected override void Execute(OleMenuCommand button)
        {
            List<ProjectItem> fls = new List<ProjectItem>();
            foreach (Project item in DTE.Solution.Projects)
                if (item.Kind == EnvDTE.Constants.vsProjectKindSolutionItems)
                    fls.AddRange(processSolutionFolder(item));
                else
                    fls.AddRange(getFiles(item.ProjectItems));

            var sBar = DTE.StatusBar;
            sBar.Text = "Очистка...";
            var toProc = fls.Where(f => f.Name.EndsWith(".cs") && !f.Name.EndsWith(".Designer.cs")).ToList();
            int i = 0;
            foreach (var file in toProc)
            {
                var window = file.Open(Constants.vsViewKindCode);
                window.Activate();
                clearCurFile();
                file.Document.Save();
                file.Document.Close();
                sBar.Progress(true, Label: "", AmountCompleted: i++, Total: toProc.Count);
            }
            sBar.Progress(false);
        }

        private static void DeleteFromBuffer(IWpfTextView view, IEnumerable<IMappingSpan> mappingSpans)
        {
            var affectedLines = new List<int>();

            RemoveCommentSpansFromBuffer(view, mappingSpans, affectedLines);
            RemoveAffectedEmptyLines(view, affectedLines);
        }

        private static void RemoveCommentSpansFromBuffer(IWpfTextView view, IEnumerable<IMappingSpan> mappingSpans, IList<int> affectedLines)
        {
            using (var edit = view.TextBuffer.CreateEdit())
            {
                foreach (var mappingSpan in mappingSpans)
                {
                    var start = mappingSpan.Start.GetPoint(view.TextBuffer, PositionAffinity.Predecessor).Value;
                    var end = mappingSpan.End.GetPoint(view.TextBuffer, PositionAffinity.Successor).Value;

                    var span = new Span(start, end - start);
                    var lines = view.TextBuffer.CurrentSnapshot.Lines.Where(l => l.Extent.IntersectsWith(span));

                    foreach (var line in lines)
                    {
                        if (IsXmlDocComment(line))
                        {
                            edit.Replace(line.Start, line.Length, string.Empty.PadLeft(line.Length));
                        }

                        if (!affectedLines.Contains(line.LineNumber))
                            affectedLines.Add(line.LineNumber);
                    }

                    var mappingText = view.TextBuffer.CurrentSnapshot.GetText(span.Start, span.Length);
                    string empty = Regex.Replace(mappingText, "([\\S]+)", string.Empty);

                    edit.Replace(span.Start, span.Length, empty);
                }

                edit.Apply();
            }
        }

        private static void RemoveAffectedEmptyLines(IWpfTextView view, IList<int> affectedLines)
        {
            if (!affectedLines.Any())
                return;

            using (var edit = view.TextBuffer.CreateEdit())
            {
                foreach (var lineNumber in affectedLines)
                {
                    var line = view.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineNumber);

                    if (IsLineEmpty(line))
                    {
                        // Strip next line if empty
                        if (view.TextBuffer.CurrentSnapshot.LineCount > line.LineNumber + 1)
                        {
                            var next = view.TextBuffer.CurrentSnapshot.GetLineFromLineNumber(lineNumber + 1);

                            if (IsLineEmpty(next))
                                edit.Delete(next.Start, next.LengthIncludingLineBreak);
                        }

                        edit.Delete(line.Start, line.LengthIncludingLineBreak);
                    }
                }

                edit.Apply();
            }
        }
    }
}
