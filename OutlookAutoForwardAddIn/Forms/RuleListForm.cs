using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using OutlookAutoForwardAddIn.Models;
using OutlookAutoForwardAddIn.Services;

namespace OutlookAutoForwardAddIn.Forms
{
    /// <summary>규칙 목록 관리 화면. 변경 사항은 [저장]을 눌러야 rules.json에 반영된다.</summary>
    public class RuleListForm : Form
    {
        private readonly ListView listView = new ListView();
        private readonly List<AutoForwardRule> rules;
        private bool loading;
        private bool dirty;

        public RuleListForm()
        {
            rules = RuleRepository.LoadRules();
            InitializeUi();
            RefreshList();
        }

        private void InitializeUi()
        {
            Text = "자동 전달 규칙 관리";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(760, 400);
            MinimumSize = new Size(640, 300);
            ShowInTaskbar = false;

            listView.View = View.Details;
            listView.CheckBoxes = true;
            listView.FullRowSelect = true;
            listView.GridLines = true;
            listView.MultiSelect = false;
            listView.HideSelection = false;
            listView.Columns.Add("사용", 70);
            listView.Columns.Add("규칙명", 200);
            listView.Columns.Add("제목 조건", 150);
            listView.Columns.Add("수신자", 280);
            listView.SetBounds(12, 12, ClientSize.Width - 24, ClientSize.Height - 68);
            listView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listView.ItemChecked += ListView_ItemChecked;
            listView.DoubleClick += (s, e) => EditSelected();
            Controls.Add(listView);

            int y = ClientSize.Height - 44;
            AddButton("추가", 12, y, (s, e) => AddRule());
            AddButton("수정", 98, y, (s, e) => EditSelected());
            AddButton("삭제", 184, y, (s, e) => DeleteSelected());

            var save = AddButton("저장", ClientSize.Width - 184, y, (s, e) => Save());
            save.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            var close = AddButton("닫기", ClientSize.Width - 98, y, (s, e) => Close());
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            CancelButton = close;

            FormClosing += RuleListForm_FormClosing;
        }

        private Button AddButton(string text, int x, int y, EventHandler onClick)
        {
            var button = new Button { Text = text, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            button.SetBounds(x, y, 80, 28);
            button.Click += onClick;
            Controls.Add(button);
            return button;
        }

        private void RefreshList()
        {
            loading = true;
            try
            {
                listView.Items.Clear();
                foreach (AutoForwardRule rule in rules)
                {
                    var item = new ListViewItem(rule.Enabled ? "Y" : "N") { Checked = rule.Enabled, Tag = rule };
                    item.SubItems.Add(rule.RuleName);
                    item.SubItems.Add(rule.SubjectContains);
                    item.SubItems.Add(rule.ToRecipients);
                    listView.Items.Add(item);
                }
            }
            finally
            {
                loading = false;
            }
        }

        private AutoForwardRule SelectedRule()
        {
            return listView.SelectedItems.Count == 0 ? null : listView.SelectedItems[0].Tag as AutoForwardRule;
        }

        private void ListView_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (loading)
                return;

            var rule = e.Item.Tag as AutoForwardRule;
            if (rule == null || rule.Enabled == e.Item.Checked)
                return;

            rule.Enabled = e.Item.Checked;
            rule.UpdatedAt = DateTime.Now;
            e.Item.Text = rule.Enabled ? "Y" : "N";
            dirty = true;
        }

        private void AddRule()
        {
            var rule = new AutoForwardRule
            {
                RuleId = RuleRepository.NewRuleId(),
                Enabled = true,
                IncludeAttachments = true,
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now
            };

            using (var form = new RuleEditForm(rule))
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                    return;
            }

            rules.Add(rule);
            dirty = true;
            RefreshList();
        }

        private void EditSelected()
        {
            AutoForwardRule selected = SelectedRule();
            if (selected == null)
            {
                MessageBox.Show(this, "수정할 규칙을 선택하세요.", "규칙 수정", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 복사본을 편집하여 취소 시 원본이 변경되지 않도록 한다.
            AutoForwardRule copy = selected.Clone();
            using (var form = new RuleEditForm(copy))
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                    return;
            }

            copy.UpdatedAt = DateTime.Now;
            rules[rules.IndexOf(selected)] = copy;
            dirty = true;
            RefreshList();
        }

        private void DeleteSelected()
        {
            AutoForwardRule selected = SelectedRule();
            if (selected == null)
            {
                MessageBox.Show(this, "삭제할 규칙을 선택하세요.", "규칙 삭제", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show(this, "규칙 '" + selected.RuleName + "'을(를) 삭제하시겠습니까?", "규칙 삭제",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            rules.Remove(selected);
            dirty = true;
            RefreshList();
        }

        private bool Save()
        {
            try
            {
                RuleRepository.SaveRules(rules);
                dirty = false;
                MessageBox.Show(this, "저장되었습니다.", "저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "규칙 저장에 실패했습니다.\n" + ex.Message, "저장 실패",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void RuleListForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!dirty)
                return;

            DialogResult answer = MessageBox.Show(this, "저장하지 않은 변경 사항이 있습니다. 저장하시겠습니까?", "닫기",
                MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (answer == DialogResult.Cancel)
                e.Cancel = true;
            else if (answer == DialogResult.Yes && !Save())
                e.Cancel = true;
        }
    }
}
