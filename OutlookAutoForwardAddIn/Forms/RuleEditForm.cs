using System;
using System.Drawing;
using System.Windows.Forms;
using OutlookAutoForwardAddIn.Models;
using OutlookAutoForwardAddIn.Utils;

namespace OutlookAutoForwardAddIn.Forms
{
    /// <summary>규칙 상세 등록/수정 화면</summary>
    public class RuleEditForm : Form
    {
        private readonly TextBox txtRuleName = new TextBox();
        private readonly TextBox txtSubject = new TextBox();
        private readonly TextBox txtExclude = new TextBox();
        private readonly TextBox txtHeader = new TextBox();
        private readonly TextBox txtTo = new TextBox();
        private readonly TextBox txtCc = new TextBox();
        private readonly TextBox txtBcc = new TextBox();
        private readonly CheckBox chkAttachments = new CheckBox();
        private readonly CheckBox chkEnabled = new CheckBox();

        public AutoForwardRule Rule { get; private set; }

        public RuleEditForm(AutoForwardRule rule)
        {
            Rule = rule;
            InitializeUi();
            LoadRule();
        }

        private void InitializeUi()
        {
            Text = "규칙 상세 설정";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 520);

            int y = 12;
            AddLabeledBox("규칙명 *", txtRuleName, ref y, 24);
            AddLabeledBox("메일 제목에 포함될 문자열 *", txtSubject, ref y, 24);
            AddLabeledBox("제목에서 제외할 단어 (금칙어, 여러 개는 ; 로 구분)", txtExclude, ref y, 24);

            txtHeader.Multiline = true;
            txtHeader.ScrollBars = ScrollBars.Vertical;
            txtHeader.AcceptsReturn = true;
            AddLabeledBox("상단에 추가할 메시지 *", txtHeader, ref y, 80);

            AddLabeledBox("수신자 * (여러 명은 ; 로 구분)", txtTo, ref y, 24);
            AddLabeledBox("참조 (여러 명은 ; 로 구분)", txtCc, ref y, 24);
            AddLabeledBox("숨은 참조 (여러 명은 ; 로 구분)", txtBcc, ref y, 24);

            chkAttachments.Text = "첨부파일 포함";
            chkAttachments.SetBounds(12, y, 200, 24);
            Controls.Add(chkAttachments);
            y += 28;

            chkEnabled.Text = "규칙 사용";
            chkEnabled.SetBounds(12, y, 200, 24);
            Controls.Add(chkEnabled);
            y += 36;

            var btnSave = new Button { Text = "저장", DialogResult = DialogResult.None };
            btnSave.SetBounds(ClientSize.Width - 176, y, 80, 28);
            btnSave.Click += BtnSave_Click;
            Controls.Add(btnSave);

            var btnCancel = new Button { Text = "취소", DialogResult = DialogResult.Cancel };
            btnCancel.SetBounds(ClientSize.Width - 92, y, 80, 28);
            Controls.Add(btnCancel);

            ClientSize = new Size(520, y + 44);
            AcceptButton = null;
            CancelButton = btnCancel;
        }

        private void AddLabeledBox(string caption, TextBox box, ref int y, int height)
        {
            var label = new Label { Text = caption, AutoSize = true };
            label.SetBounds(12, y, 480, 18);
            Controls.Add(label);
            y += 20;

            box.SetBounds(12, y, ClientSize.Width - 24, height);
            Controls.Add(box);
            y += height + 10;
        }

        private void LoadRule()
        {
            txtRuleName.Text = Rule.RuleName;
            txtSubject.Text = Rule.SubjectContains;
            txtExclude.Text = Rule.ExcludeSubjects;
            txtHeader.Text = Rule.HeaderMessage;
            txtTo.Text = Rule.ToRecipients;
            txtCc.Text = Rule.CcRecipients;
            txtBcc.Text = Rule.BccRecipients;
            chkAttachments.Checked = Rule.IncludeAttachments;
            chkEnabled.Checked = Rule.Enabled;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!Validate(out string error, out Control focus))
            {
                MessageBox.Show(this, error, "입력 확인", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                focus.Focus();
                return;
            }

            Rule.RuleName = txtRuleName.Text.Trim();
            Rule.SubjectContains = txtSubject.Text.Trim();
            Rule.ExcludeSubjects = EmailValidator.Normalize(txtExclude.Text);
            Rule.HeaderMessage = txtHeader.Text.Trim();
            Rule.ToRecipients = EmailValidator.Normalize(txtTo.Text);
            Rule.CcRecipients = EmailValidator.Normalize(txtCc.Text);
            Rule.BccRecipients = EmailValidator.Normalize(txtBcc.Text);
            Rule.IncludeAttachments = chkAttachments.Checked;
            Rule.Enabled = chkEnabled.Checked;

            DialogResult = DialogResult.OK;
            Close();
        }

        private bool Validate(out string error, out Control focus)
        {
            error = null;
            focus = null;

            if (string.IsNullOrWhiteSpace(txtRuleName.Text))
                return Fail("규칙명을 입력하세요.", txtRuleName, out error, out focus);
            if (string.IsNullOrWhiteSpace(txtSubject.Text))
                return Fail("제목 포함 문자열을 입력하세요.", txtSubject, out error, out focus);
            if (string.IsNullOrWhiteSpace(txtHeader.Text))
                return Fail("상단에 추가할 메시지를 입력하세요.", txtHeader, out error, out focus);
            if (EmailValidator.Split(txtTo.Text).Count == 0)
                return Fail("수신자를 입력하세요.", txtTo, out error, out focus);

            string invalid;
            if (!EmailValidator.TryGetInvalid(txtTo.Text, out invalid))
                return Fail("수신자 이메일 주소가 올바르지 않습니다: " + invalid, txtTo, out error, out focus);
            if (!EmailValidator.TryGetInvalid(txtCc.Text, out invalid))
                return Fail("참조자 이메일 주소가 올바르지 않습니다: " + invalid, txtCc, out error, out focus);
            if (!EmailValidator.TryGetInvalid(txtBcc.Text, out invalid))
                return Fail("숨은 참조자 이메일 주소가 올바르지 않습니다: " + invalid, txtBcc, out error, out focus);

            return true;
        }

        private static bool Fail(string message, Control control, out string error, out Control focus)
        {
            error = message;
            focus = control;
            return false;
        }
    }
}
