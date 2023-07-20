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
        static readonly string[] exceptDir = { "obj", "debug", "bin" };
        /// <summary>
        /// Рекурсивный метод извлечения перечня файлов
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        internal IEnumerable<String> getDirFiles(string dir)
        {
            List<String> files = new List<String>();
            files.AddRange(Directory.GetFiles(dir, "*.cs"));
            foreach (var sdir in Directory.GetDirectories(dir))
                if (!exceptDir.Contains(Path.GetFileName(sdir).ToLower()))
                    files.AddRange(getDirFiles(sdir));
            return files;
        }

        internal void clearCurFile()
        {
            var view = ProjectHelpers.GetCurentTextView();
            var mappingSpans = GetClassificationSpans(view, "comment");

            if (!mappingSpans.Any())
                return;

            try
            {
                //DTE.UndoContext.Open(button.Text);

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

        //private const String SfGuid = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
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
            var toProc = fls.Where(f => f.Name.EndsWith(".cs")).ToList();
            int i = 0;
            foreach (var file in toProc)
            {
                //var window = DTE.ItemOperations.OpenFile(allFiles[i]);
                var window = file.Open();
                window.Activate();
                //DTE.ExecuteCommand("File.OpenFile", );
                clearCurFile();
                file.Document.Save();
                file.Document.Close();
                //DTE.ExecuteCommand("File.Save");
                //window.Close();
                //DTE.ExecuteCommand("File.Close");
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
