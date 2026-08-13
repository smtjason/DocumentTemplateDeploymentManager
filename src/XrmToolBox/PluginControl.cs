using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;

namespace XrmToolBox.DocumentTemplateDeploymentManager
{
    public sealed class PluginControl : MultipleConnectionsPluginControlBase
    {
        private readonly DocumentTemplateDeploymentService deploymentService = new DocumentTemplateDeploymentService();
        private readonly ToolStrip toolStrip = new ToolStrip();
        private readonly ToolStripButton loadButton = new ToolStripButton("Load Source Templates");
        private readonly ToolStripButton addTargetButton = new ToolStripButton("Add Target Connection");
        private readonly ToolStripButton removeTargetButton = new ToolStripButton("Remove Selected Target");
        private readonly ToolStripButton moveButton = new ToolStripButton("Move Selected Templates");
        private readonly System.Windows.Forms.Label sourceLabel = new System.Windows.Forms.Label();
        private readonly ListView templateList = new ListView();
        private readonly CheckedListBox targetList = new CheckedListBox();
        private readonly DataGridView resultGrid = new DataGridView();
        private readonly BindingList<MoveResult> results = new BindingList<MoveResult>();
        private ConnectionDetail sourceConnectionDetail;
        private readonly string[] templateColumnTitles = { "Name", "Created", "Modified", "Associated table" };
        private int templateSortColumn = -1;
        private SortOrder templateSortOrder = SortOrder.None;

        public PluginControl()
        {
            BuildUi();
            loadButton.Click += (_, __) => ExecuteMethod(LoadTemplates);
            addTargetButton.Click += (_, __) => AddAdditionalOrganization();
            removeTargetButton.Click += (_, __) => RemoveSelectedTarget();
            moveButton.Click += (_, __) => MoveSelectedTemplates();
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName = "", object parameter = null)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            if (!string.Equals(actionName, "AdditionalOrganization", StringComparison.OrdinalIgnoreCase))
            {
                sourceConnectionDetail = detail;
                sourceLabel.Text = "Source: " + (detail?.ConnectionName ?? "Not connected");
                sourceLabel.ForeColor = detail == null ? Color.DarkRed : Color.DarkGreen;
                RefreshTargets();
            }
            else if (IsSourceConnection(detail))
            {
                BeginInvoke(new Action(() =>
                {
                    RemoveAdditionalOrganization(detail);
                    MessageBox.Show(this, "The Source environment cannot also be selected as a Target.", "Invalid Target", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }));
            }
        }

        protected override void ConnectionDetailsUpdated(NotifyCollectionChangedEventArgs e) => RefreshTargets();

        private void BuildUi()
        {
            Dock = DockStyle.Fill;
            toolStrip.GripStyle = ToolStripGripStyle.Hidden;
            toolStrip.Items.AddRange(new ToolStripItem[] { loadButton, new ToolStripSeparator(), addTargetButton, removeTargetButton, new ToolStripSeparator(), moveButton });
            toolStrip.Dock = DockStyle.Top;

            sourceLabel.Text = "Source: select the main XrmToolBox connection";
            sourceLabel.AutoSize = false;
            sourceLabel.Height = 30;
            sourceLabel.Padding = new Padding(8, 7, 0, 0);
            sourceLabel.Dock = DockStyle.Top;

            templateList.CheckBoxes = true;
            templateList.FullRowSelect = true;
            templateList.GridLines = true;
            templateList.HideSelection = false;
            templateList.View = View.Details;
            templateList.Columns.Add("Name", 330);
            templateList.Columns.Add("Created", 145);
            templateList.Columns.Add("Modified", 145);
            templateList.Columns.Add("Associated table", 260);
            templateList.ColumnClick += TemplateListColumnClick;
            templateList.Dock = DockStyle.Fill;

            targetList.CheckOnClick = true;
            targetList.Dock = DockStyle.Fill;

            resultGrid.AutoGenerateColumns = false;
            resultGrid.AllowUserToAddRows = false;
            resultGrid.AllowUserToDeleteRows = false;
            resultGrid.ReadOnly = true;
            resultGrid.RowHeadersVisible = false;
            resultGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Target", DataPropertyName = "Target", FillWeight = 20 });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Template", DataPropertyName = "Template", FillWeight = 25 });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Action", DataPropertyName = "Action", FillWeight = 10 });
            resultGrid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Success", DataPropertyName = "Success", FillWeight = 8 });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Bindings", DataPropertyName = "Bindings", FillWeight = 8 });
            resultGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Result / error", DataPropertyName = "Message", FillWeight = 40 });
            resultGrid.DataSource = results;
            resultGrid.Dock = DockStyle.Fill;
            resultGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            resultGrid.KeyDown += ResultGridKeyDown;
            var resultsMenu = new ContextMenuStrip();
            resultsMenu.Items.Add("Copy all deployment results", null, (_, __) => CopyDeploymentResults());
            resultGrid.ContextMenuStrip = resultsMenu;

            var sourceGroup = new GroupBox { Text = "Source Word document templates", Dock = DockStyle.Fill };
            sourceGroup.Controls.Add(templateList);
            var targetGroup = new GroupBox { Text = "Target connections (check one or more)", Dock = DockStyle.Fill };
            targetGroup.Controls.Add(targetList);
            var upperSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 670 };
            upperSplit.Panel1.Controls.Add(sourceGroup);
            upperSplit.Panel2.Controls.Add(targetGroup);
            var resultsGroup = new GroupBox { Text = "Deployment results", Dock = DockStyle.Fill };
            resultsGroup.Controls.Add(resultGrid);
            var mainSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 360 };
            mainSplit.Panel1.Controls.Add(upperSplit);
            mainSplit.Panel2.Controls.Add(resultsGroup);
            Controls.Add(mainSplit);
            Controls.Add(sourceLabel);
            Controls.Add(toolStrip);
        }

        private void LoadTemplates()
        {
            if (Service == null) return;
            SetBusy(true);
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading Word document templates from Source...",
                Work = (_, e) => e.Result = deploymentService.GetWordTemplates(Service),
                PostWorkCallBack = e =>
                {
                    SetBusy(false);
                    if (e.Error != null) { ShowError("Could not load templates", e.Error); return; }
                    templateList.Items.Clear();
                    foreach (var template in (IReadOnlyList<TemplateItem>)e.Result)
                    {
                        var item = new ListViewItem(template.Name) { Tag = template };
                        item.SubItems.Add(FormatDate(template.CreatedOn));
                        item.SubItems.Add(FormatDate(template.ModifiedOn));
                        item.SubItems.Add(template.AssociatedTable);
                        templateList.Items.Add(item);
                    }
                    ApplyTemplateSort();
                }
            });
        }

        private void MoveSelectedTemplates()
        {
            var templates = templateList.CheckedItems.Cast<ListViewItem>().Select(i => (TemplateItem)i.Tag).ToList();
            var targets = targetList.CheckedItems.Cast<TargetItem>().ToList();
            if (targets.Any(target => IsSourceConnection(target.Detail)))
            {
                MessageBox.Show(this, "The Source environment cannot also be selected as a Target. Remove it from the Target list before deploying.", "Invalid Target", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RefreshTargets();
                return;
            }
            if (templates.Count == 0 || targets.Count == 0)
            {
                MessageBox.Show(this, "Select at least one Source template and one Target connection.", "Selection required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var prompt = $"Move {templates.Count} template(s) to {targets.Count} target environment(s)?\n\nExisting exact matches will be updated; missing templates will be created.";
            if (MessageBox.Show(this, prompt, "Confirm document-template deployment", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;

            results.Clear();
            SetBusy(true);
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Moving and validating Word document templates...",
                Work = (_, e) =>
                {
                    var completed = new List<MoveResult>();
                    foreach (var target in targets)
                    {
                        var service = target.Detail.GetCrmServiceClient();
                        foreach (var template in templates)
                        {
                            try { completed.Add(deploymentService.Deploy(service, target.ToString(), template.Record)); }
                            catch (Exception ex)
                            {
                                completed.Add(new MoveResult { Target = target.ToString(), Template = template.Name, Action = "Failed", Success = false, Bindings = 0, Message = ex.Message });
                            }
                        }
                    }
                    e.Result = completed;
                },
                PostWorkCallBack = e =>
                {
                    SetBusy(false);
                    if (e.Error != null) { ShowError("Deployment stopped unexpectedly", e.Error); return; }
                    foreach (var result in (IReadOnlyList<MoveResult>)e.Result) results.Add(result);
                    var failures = results.Count(r => !r.Success);
                    MessageBox.Show(this,
                        failures == 0 ? "All template deployments completed and validated successfully." : $"Deployment completed with {failures} failure(s). Review the results grid.",
                        "Deployment complete", MessageBoxButtons.OK,
                        failures == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                }
            });
        }

        private void RefreshTargets()
        {
            var checkedConnections = targetList.CheckedItems.Cast<TargetItem>().Select(i => i.Detail.ConnectionId).ToHashSet();
            targetList.Items.Clear();
            foreach (var detail in AdditionalConnectionDetails.Where(detail => !IsSourceConnection(detail)))
            {
                var item = new TargetItem { Detail = detail };
                targetList.Items.Add(item, checkedConnections.Contains(detail.ConnectionId) || checkedConnections.Count == 0);
            }
        }

        private bool IsSourceConnection(ConnectionDetail detail)
        {
            if (detail == null || sourceConnectionDetail == null) return false;
            if (detail.ConnectionId != Guid.Empty && detail.ConnectionId == sourceConnectionDetail.ConnectionId) return true;
            if (!string.IsNullOrWhiteSpace(detail.EnvironmentId) && string.Equals(detail.EnvironmentId, sourceConnectionDetail.EnvironmentId, StringComparison.OrdinalIgnoreCase)) return true;
            return SameUrl(detail.OrganizationServiceUrl, sourceConnectionDetail.OrganizationServiceUrl) ||
                   SameUrl(detail.WebApplicationUrl, sourceConnectionDetail.WebApplicationUrl);
        }

        private static bool SameUrl(string left, string right)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
            return string.Equals(left.Trim().TrimEnd('/'), right.Trim().TrimEnd('/'), StringComparison.OrdinalIgnoreCase);
        }

        private void RemoveSelectedTarget()
        {
            if (targetList.SelectedItem is TargetItem item) RemoveAdditionalOrganization(item.Detail);
        }

        private void TemplateListColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (templateSortColumn == e.Column)
                templateSortOrder = templateSortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            else
            {
                templateSortColumn = e.Column;
                templateSortOrder = SortOrder.Ascending;
            }

            ApplyTemplateSort();
        }

        private void ApplyTemplateSort()
        {
            for (var index = 0; index < templateColumnTitles.Length; index++)
                templateList.Columns[index].Text = templateColumnTitles[index];

            if (templateSortColumn < 0) return;
            templateList.Columns[templateSortColumn].Text += templateSortOrder == SortOrder.Ascending ? " ▲" : " ▼";
            templateList.ListViewItemSorter = new TemplateListComparer(templateSortColumn, templateSortOrder);
            templateList.Sort();
        }

        private void CopyDeploymentResults()
        {
            if (results.Count == 0)
            {
                MessageBox.Show(this, "There are no deployment results to copy.", "Nothing to copy", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var markdown = new StringBuilder();
            markdown.AppendLine("**Document template deployment results**");
            markdown.AppendLine();
            markdown.AppendLine("| Target | Template | Action | Success | Bindings | Result / error |");
            markdown.AppendLine("|---|---|---|:---:|---:|---|");

            var htmlRows = new StringBuilder();
            foreach (var result in results)
            {
                markdown.Append("| ").Append(EscapeMarkdown(result.Target)).Append(" | ")
                    .Append(EscapeMarkdown(result.Template)).Append(" | ")
                    .Append(EscapeMarkdown(result.Action)).Append(" | ")
                    .Append(result.Success ? "Yes" : "No").Append(" | ")
                    .Append(result.Bindings).Append(" | ")
                    .Append(EscapeMarkdown(result.Message)).AppendLine(" |");

                var statusColor = result.Success ? "#107c10" : "#a4262c";
                htmlRows.Append("<tr>")
                    .Append(HtmlCell(result.Target)).Append(HtmlCell(result.Template)).Append(HtmlCell(result.Action))
                    .Append("<td style=\"padding:7px 10px;border:1px solid #d1d1d1;color:").Append(statusColor).Append(";font-weight:600\">")
                    .Append(result.Success ? "Yes" : "No").Append("</td>")
                    .Append(HtmlCell(result.Bindings.ToString())).Append(HtmlCell(result.Message)).Append("</tr>");
            }

            var fragment = "<div style=\"font-family:Segoe UI,Arial,sans-serif;font-size:10pt\">" +
                "<div style=\"font-size:14pt;font-weight:600;margin-bottom:8px\">Document template deployment results</div>" +
                "<table style=\"border-collapse:collapse\"><thead><tr style=\"background:#0078d4;color:white\">" +
                HtmlHeader("Target") + HtmlHeader("Template") + HtmlHeader("Action") + HtmlHeader("Success") + HtmlHeader("Bindings") + HtmlHeader("Result / error") +
                "</tr></thead><tbody>" + htmlRows + "</tbody></table></div>";

            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, markdown.ToString());
            data.SetData(DataFormats.Text, markdown.ToString());
            data.SetData(DataFormats.Html, BuildClipboardHtml(fragment));
            Clipboard.SetDataObject(data, true);
        }

        private void ResultGridKeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.C) return;
            CopyDeploymentResults();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private static string CleanClipboardValue(string value) => (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
        private static string EscapeMarkdown(string value) => CleanClipboardValue(value).Replace("|", "\\|");
        private static string HtmlCell(string value) => "<td style=\"padding:7px 10px;border:1px solid #d1d1d1;vertical-align:top\">" + System.Net.WebUtility.HtmlEncode(CleanClipboardValue(value)) + "</td>";
        private static string HtmlHeader(string value) => "<th style=\"padding:8px 10px;border:1px solid #d1d1d1;text-align:left\">" + System.Net.WebUtility.HtmlEncode(value) + "</th>";

        private static string BuildClipboardHtml(string fragment)
        {
            const string headerTemplate = "Version:0.9\r\nStartHTML:{0:0000000000}\r\nEndHTML:{1:0000000000}\r\nStartFragment:{2:0000000000}\r\nEndFragment:{3:0000000000}\r\n";
            const string startMarker = "<!--StartFragment-->";
            const string endMarker = "<!--EndFragment-->";
            var html = "<html><body>" + startMarker + fragment + endMarker + "</body></html>";
            var placeholderHeader = string.Format(headerTemplate, 0, 0, 0, 0);
            var startHtml = Encoding.UTF8.GetByteCount(placeholderHeader);
            var startFragment = startHtml + Encoding.UTF8.GetByteCount("<html><body>" + startMarker);
            var endFragment = startFragment + Encoding.UTF8.GetByteCount(fragment);
            var endHtml = startHtml + Encoding.UTF8.GetByteCount(html);
            return string.Format(headerTemplate, startHtml, endHtml, startFragment, endFragment) + html;
        }

        private void SetBusy(bool busy)
        {
            loadButton.Enabled = !busy;
            addTargetButton.Enabled = !busy;
            removeTargetButton.Enabled = !busy;
            moveButton.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        private static string FormatDate(DateTime? value) => value.HasValue ? value.Value.ToLocalTime().ToString("g") : string.Empty;
        private void ShowError(string title, Exception error) => MessageBox.Show(this, error.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        private sealed class TemplateListComparer : System.Collections.IComparer
        {
            private readonly int column;
            private readonly int direction;

            public TemplateListComparer(int column, SortOrder order)
            {
                this.column = column;
                direction = order == SortOrder.Descending ? -1 : 1;
            }

            public int Compare(object x, object y)
            {
                var left = ((ListViewItem)x).Tag as TemplateItem;
                var right = ((ListViewItem)y).Tag as TemplateItem;
                int comparison;
                switch (column)
                {
                    case 1:
                        comparison = Nullable.Compare(left?.CreatedOn, right?.CreatedOn);
                        break;
                    case 2:
                        comparison = Nullable.Compare(left?.ModifiedOn, right?.ModifiedOn);
                        break;
                    case 3:
                        comparison = StringComparer.CurrentCultureIgnoreCase.Compare(left?.AssociatedTable, right?.AssociatedTable);
                        break;
                    default:
                        comparison = StringComparer.CurrentCultureIgnoreCase.Compare(left?.Name, right?.Name);
                        break;
                }
                return comparison * direction;
            }
        }
    }
}
