# Outlook VSTO Add-in 개발 계획서

## 1. 개발 목적

Microsoft Outlook Desktop 환경에서 특정 제목 조건에 맞는 메일이 수신되었을 때, 사전에 사용자가 설정해 둔 메시지를 메일 본문 상단에 자동으로 추가하고, 사전에 지정한 수신자에게 자동 전달하는 VSTO Add-in을 개발한다.

본 기능은 Outlook 기본 규칙과 유사하게 동작하지만, 전달 메일 상단에 사용자가 사전에 정의한 안내 메시지를 자동 삽입할 수 있다는 점이 차별점이다.

---

## 2. 개발 대상

| 항목 | 내용 |
|---|---|
| 대상 애플리케이션 | Microsoft Outlook Desktop |
| 개발 방식 | VSTO Add-in |
| 개발 언어 | C# |
| 개발 도구 | Visual Studio |
| 대상 플랫폼 | Windows |
| 주요 API | Microsoft.Office.Interop.Outlook |

---

## 3. 핵심 요구사항

## 3.1 규칙 등록 기능

사용자는 Outlook Add-in 내 설정 화면에서 자동 전달 규칙을 등록할 수 있어야 한다.

규칙에는 다음 항목이 포함된다.

| 항목 | 필수 여부 | 설명 |
|---|---:|---|
| 규칙명 | 필수 | 사용자가 규칙을 식별하기 위한 이름 |
| 제목 포함 문자열 | 필수 | 수신 메일 제목에 포함되어야 하는 문자열 |
| 상단 추가 메시지 | 필수 | 전달 메일 본문 상단에 삽입할 메시지 |
| 수신자 | 필수 | 자동 전달 대상 이메일 주소 |
| 참조자 | 선택 | 자동 전달 시 참조 이메일 주소 |
| 숨은 참조자 | 선택 | 자동 전달 시 숨은 참조 이메일 주소 |
| 규칙 사용 여부 | 필수 | 해당 규칙 활성화 여부 |
| 첨부파일 포함 여부 | 선택 | 원본 메일 첨부파일 포함 여부 |

---

## 3.2 메일 수신 감지 기능

Outlook에서 신규 메일이 수신되면 Add-in이 이를 감지해야 한다.

권장 방식은 다음 중 하나를 사용한다.

| 방식 | 설명 |
|---|---|
| Application.NewMailEx | 신규 메일의 EntryID를 기반으로 메일을 조회 |
| Items.ItemAdd | 수신함 폴더의 신규 아이템 추가 이벤트 감지 |

초기 구현은 `Application.NewMailEx`를 우선 사용한다.

---

## 3.3 규칙 매칭 기능

수신 메일의 제목을 확인하여 등록된 규칙의 `제목 포함 문자열`과 비교한다.

예를 들어 다음 규칙이 등록되어 있다고 가정한다.

| 규칙명 | 제목 포함 문자열 |
|---|---|
| 설비 점검 자동 전달 | 설비 점검 |

수신 메일 제목이 `[긴급] 설비 점검 요청`인 경우, 제목에 `설비 점검`이 포함되어 있으므로 해당 규칙이 실행되어야 한다.

---

## 3.4 자동 전달 기능

조건에 맞는 메일이 수신되면 다음 순서로 자동 전달한다.

1. 원본 메일을 기준으로 전달 메일을 생성한다.
2. 규칙에 저장된 수신자, 참조자, 숨은 참조자를 설정한다.
3. 규칙에 저장된 상단 추가 메시지를 전달 메일 본문 최상단에 삽입한다.
4. 원본 메일 내용을 하단에 유지한다.
5. 설정에 따라 첨부파일 포함 여부를 처리한다.
6. 메일을 자동 발송한다.

---

## 4. 기본 사용 흐름

## 4.1 규칙 등록 흐름

1. 사용자가 Outlook을 실행한다.
2. VSTO Add-in이 Outlook에 로드된다.
3. 사용자가 Add-in의 규칙 관리 화면을 연다.
4. 사용자가 자동 전달 규칙을 신규 등록한다.
5. 사용자는 다음 정보를 입력한다.
   - 규칙명
   - 제목 포함 문자열
   - 상단 추가 메시지
   - 수신자
   - 참조자
   - 숨은 참조자
   - 첨부파일 포함 여부
   - 규칙 사용 여부
6. 사용자가 저장 버튼을 클릭한다.
7. Add-in은 규칙 정보를 로컬 설정 파일에 저장한다.

---

## 4.2 자동 전달 흐름

1. Outlook에 신규 메일이 수신된다.
2. Add-in이 신규 메일 수신 이벤트를 감지한다.
3. Add-in이 수신 메일의 제목을 확인한다.
4. 저장된 규칙 목록을 불러온다.
5. 사용 상태인 규칙만 대상으로 제목 조건을 비교한다.
6. 제목 조건과 일치하는 규칙이 있으면 전달 메일을 생성한다.
7. 규칙에 저장된 상단 추가 메시지를 전달 메일 본문 상단에 삽입한다.
8. 규칙에 저장된 수신자, 참조자, 숨은 참조자를 설정한다.
9. 메일을 자동 발송한다.

---

## 5. 기능 예시

## 5.1 등록된 규칙 예시

| 항목 | 설정값 |
|---|---|
| 규칙명 | 설비 점검 자동 전달 |
| 제목 포함 문자열 | 설비 점검 |
| 상단 추가 메시지 | 아래 요청 내용 확인 후 금일 중 조치 부탁드립니다. |
| 수신자 | maintenance@company.com |
| 참조자 | manager@company.com |
| 숨은 참조자 | |
| 첨부파일 포함 여부 | 포함 |
| 규칙 사용 여부 | 사용 |

---

## 5.2 수신 메일 예시

| 항목 | 내용 |
|---|---|
| 보낸 사람 | request@company.com |
| 제목 | [긴급] 설비 점검 요청 |
| 본문 | 1라인 설비 점검 요청드립니다. |

---

## 5.3 자동 전달 결과 예시

### 수신자

```text
maintenance@company.com
```

### 참조자

```text
manager@company.com
```

### 본문 예시

```text
아래 요청 내용 확인 후 금일 중 조치 부탁드립니다.

----- 원본 메일 -----

보낸 사람: request@company.com
제목: [긴급] 설비 점검 요청

1라인 설비 점검 요청드립니다.
```

---

## 6. 화면 요구사항

## 6.1 규칙 목록 화면

규칙 목록 화면에서는 사용자가 등록된 규칙을 확인하고 관리할 수 있어야 한다.

### 필요 기능

- 규칙 목록 조회
- 규칙 추가
- 규칙 수정
- 규칙 삭제
- 규칙 사용 여부 변경
- 저장
- 닫기

### 화면 예시

```text
[자동 전달 규칙 관리]

------------------------------------------------------------
사용 | 규칙명              | 제목 조건     | 수신자
------------------------------------------------------------
 Y   | 설비 점검 자동 전달 | 설비 점검     | maintenance@company.com
 Y   | 보안 알림 자동 전달 | 보안 알림     | security@company.com
 N   | 테스트 규칙          | TEST          | test@company.com
------------------------------------------------------------

[추가] [수정] [삭제] [저장] [닫기]
```

---

## 6.2 규칙 상세 화면

규칙 상세 화면에서는 사용자가 자동 전달 조건과 메시지, 수신자를 설정할 수 있어야 한다.

### 입력 항목

| 항목 | 컨트롤 유형 | 필수 여부 |
|---|---|---:|
| 규칙명 | TextBox | 필수 |
| 제목 포함 문자열 | TextBox | 필수 |
| 상단 추가 메시지 | Multi-line TextBox | 필수 |
| 수신자 | TextBox | 필수 |
| 참조자 | TextBox | 선택 |
| 숨은 참조자 | TextBox | 선택 |
| 첨부파일 포함 여부 | CheckBox | 선택 |
| 규칙 사용 여부 | CheckBox | 필수 |

### 화면 예시

```text
[규칙 상세 설정]

규칙명:
[설비 점검 자동 전달]

메일 제목에 포함될 문자열:
[설비 점검]

상단에 추가할 메시지:
[아래 요청 내용 확인 후 금일 중 조치 부탁드립니다.]

수신자:
[maintenance@company.com]

참조:
[manager@company.com]

숨은 참조:
[]

첨부파일 포함:
[✓]

규칙 사용:
[✓]

[저장] [취소]
```

---

## 7. 데이터 저장 방식

초기 버전에서는 규칙 정보를 JSON 파일로 저장한다.

저장 위치는 사용자별 설정 경로를 사용한다.

예시 경로:

```text
%AppData%\OutlookAutoForwardAddIn\rules.json
```

---

## 8. 데이터 모델

## 8.1 AutoForwardRule 모델

```csharp
public class AutoForwardRule
{
    public string RuleId { get; set; }
    public string RuleName { get; set; }
    public bool Enabled { get; set; }

    public string SubjectContains { get; set; }
    public string HeaderMessage { get; set; }

    public string ToRecipients { get; set; }
    public string CcRecipients { get; set; }
    public string BccRecipients { get; set; }

    public bool IncludeAttachments { get; set; }

    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }
}
```

---

## 8.2 JSON 저장 예시

```json
[
  {
    "RuleId": "RULE-001",
    "RuleName": "설비 점검 자동 전달",
    "Enabled": true,
    "SubjectContains": "설비 점검",
    "HeaderMessage": "아래 요청 내용 확인 후 금일 중 조치 부탁드립니다.",
    "ToRecipients": "maintenance@company.com",
    "CcRecipients": "manager@company.com",
    "BccRecipients": "",
    "IncludeAttachments": true,
    "CreatedAt": "2026-10-19T09:00:00",
    "UpdatedAt": "2026-10-19T09:00:00"
  }
]
```

---

## 9. 프로젝트 구조

```text
OutlookAutoForwardAddIn
├─ ThisAddIn.cs
├─ Models
│  └─ AutoForwardRule.cs
├─ Services
│  ├─ MailMonitorService.cs
│  ├─ RuleRepository.cs
│  ├─ RuleMatchService.cs
│  └─ MailForwardService.cs
├─ Forms
│  ├─ RuleListForm.cs
│  └─ RuleEditForm.cs
└─ Utils
   └─ HtmlEncodeHelper.cs
```

---

## 10. 주요 클래스 역할

| 클래스 | 역할 |
|---|---|
| ThisAddIn | Outlook Add-in 시작점, 이벤트 연결 |
| MailMonitorService | 신규 메일 수신 감지 |
| RuleRepository | 규칙 JSON 저장 및 로드 |
| RuleMatchService | 수신 메일과 규칙 조건 비교 |
| MailForwardService | 전달 메일 생성, 메시지 삽입, 자동 발송 |
| RuleListForm | 규칙 목록 관리 화면 |
| RuleEditForm | 규칙 상세 등록 및 수정 화면 |
| HtmlEncodeHelper | HTML 본문 삽입 시 특수문자 인코딩 처리 |

---

## 11. 구현 상세

## 11.1 Outlook 시작 시 이벤트 연결

```csharp
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    this.Application.NewMailEx += Application_NewMailEx;
}
```

---

## 11.2 신규 메일 수신 처리

```csharp
private void Application_NewMailEx(string entryIDCollection)
{
    try
    {
        var mail = Application.Session.GetItemFromID(entryIDCollection) as Outlook.MailItem;

        if (mail == null)
            return;

        if (string.IsNullOrWhiteSpace(mail.Subject))
            return;

        var rules = RuleRepository.LoadRules();

        var matchedRule = rules
            .Where(r => r.Enabled)
            .FirstOrDefault(r => mail.Subject.Contains(r.SubjectContains));

        if (matchedRule == null)
            return;

        MailForwardService.ForwardByRule(mail, matchedRule);
    }
    catch
    {
        // 초기 버전에서는 별도 로그 기능을 구현하지 않는다.
        // 필요한 경우 예외 메시지를 사용자에게 표시하거나 디버그 출력으로만 처리한다.
    }
}
```

---

## 11.3 규칙 기반 자동 전달

```csharp
public void ForwardByRule(Outlook.MailItem originalMail, AutoForwardRule rule)
{
    Outlook.MailItem forwardMail = originalMail.Forward();

    forwardMail.To = rule.ToRecipients;
    forwardMail.CC = rule.CcRecipients;
    forwardMail.BCC = rule.BccRecipients;

    if (originalMail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
    {
        string encodedMessage = HtmlEncode(rule.HeaderMessage)
            .Replace("\r\n", "<br/>")
            .Replace("\n", "<br/>");

        forwardMail.HTMLBody =
            "<div style='font-family:Malgun Gothic; font-size:10pt;'>" +
            encodedMessage +
            "</div><hr/>" +
            forwardMail.HTMLBody;
    }
    else
    {
        forwardMail.Body =
            rule.HeaderMessage +
            Environment.NewLine +
            Environment.NewLine +
            "----- 원본 메일 -----" +
            Environment.NewLine +
            forwardMail.Body;
    }

    if (!rule.IncludeAttachments)
    {
        RemoveAttachments(forwardMail);
    }

    forwardMail.Send();
}
```

---

## 11.4 첨부파일 제거 처리

```csharp
private void RemoveAttachments(Outlook.MailItem mail)
{
    for (int i = mail.Attachments.Count; i >= 1; i--)
    {
        mail.Attachments.Remove(i);
    }
}
```

---

## 11.5 HTML 인코딩 처리

```csharp
private string HtmlEncode(string value)
{
    if (string.IsNullOrEmpty(value))
        return string.Empty;

    return System.Net.WebUtility.HtmlEncode(value);
}
```

---

## 12. 여러 규칙이 동시에 매칭되는 경우

초기 버전에서는 여러 규칙이 동시에 매칭되더라도 첫 번째로 매칭된 규칙만 실행한다.

### 사유

- 동일 메일의 중복 발송 방지
- 오발송 가능성 감소
- 기능 단순화
- 운영 관리 용이성 확보

향후 필요 시 규칙 우선순위 또는 모든 매칭 규칙 실행 옵션을 추가할 수 있다.

---

## 13. 예외 처리 방안

| 상황 | 처리 방안 |
|---|---|
| 규칙 파일이 없음 | 빈 규칙 목록으로 처리 |
| 규칙 JSON 파싱 실패 | 규칙 목록을 비운 상태로 처리 |
| 제목이 없는 메일 | 처리하지 않음 |
| 규칙이 비활성화 상태 | 처리하지 않음 |
| 수신자 값이 비어 있음 | 해당 규칙 저장 불가 |
| 제목 조건 값이 비어 있음 | 해당 규칙 저장 불가 |
| 상단 메시지가 비어 있음 | 해당 규칙 저장 불가 |
| 자동 발송 실패 | 예외 처리 후 Outlook 동작에 영향이 없도록 함 |
| HTML 본문 처리 실패 | 일반 텍스트 방식으로 대체 가능 |

---

## 14. 보안 및 운영 고려사항

본 기능은 조건에 맞는 메일을 사용자 확인 없이 자동 발송하므로 다음 사항을 고려해야 한다.

- 수신자 이메일 주소 유효성 검사를 수행한다.
- 외부 도메인 발송 제한이 필요한지 검토한다.
- 첨부파일 포함 여부를 규칙별로 선택할 수 있게 한다.
- 동일 메일이 중복으로 처리되지 않도록 보완이 필요할 수 있다.
- 조직의 Outlook 보안 정책에 따라 자동 발송 시 보안 경고가 표시될 수 있다.
- 사용자가 규칙을 명확하게 관리할 수 있도록 규칙 사용 여부 설정을 제공한다.

---

## 15. 중복 발송 방지

초기 버전에서는 간단 구현을 우선하되, 가능하면 동일 메일 중복 발송 방지를 구현한다.

권장 기준은 다음 중 하나를 사용한다.

- EntryID
- InternetMessageId
- 수신 시간 + 제목 + 보낸 사람 조합

단, 이번 버전에서는 로그 기능은 구현하지 않는다.

중복 처리 정보가 필요한 경우 메모리 기반 임시 목록을 사용할 수 있다.

예시:

```csharp
private readonly HashSet<string> processedEntryIds = new HashSet<string>();
```

---

## 16. 제외 범위

이번 개발 범위에서 제외할 기능은 다음과 같다.

- 처리 로그 저장 기능
- 로그 조회 화면
- 관리자 승인 기능
- Exchange 서버 측 규칙 연동
- Microsoft Graph API 연동
- 다중 사용자 중앙 관리 기능
- 규칙별 우선순위 상세 설정
- 예약 발송 기능
- 메일 본문 키워드 조건
- 보낸 사람 조건
- 첨부파일명 조건

---

## 17. 개발 일정

| 단계 | 작업 내용 | 예상 기간 |
|---|---|---:|
| 1 | VSTO Add-in 프로젝트 생성 | 0.5일 |
| 2 | Outlook 신규 메일 수신 이벤트 연결 | 0.5일 |
| 3 | 규칙 데이터 모델 구현 | 0.5일 |
| 4 | JSON 기반 규칙 저장 및 로드 구현 | 1일 |
| 5 | 규칙 목록 화면 구현 | 1일 |
| 6 | 규칙 상세 등록 및 수정 화면 구현 | 1일 |
| 7 | 제목 조건 매칭 로직 구현 | 0.5일 |
| 8 | 전달 메일 생성 및 상단 메시지 삽입 구현 | 1일 |
| 9 | 자동 발송 기능 구현 | 0.5일 |
| 10 | 첨부파일 포함 여부 처리 | 0.5일 |
| 11 | 예외 처리 및 테스트 | 1~2일 |


---

## 18. 테스트 항목

## 18.1 규칙 관리 테스트

| 테스트 항목 | 기대 결과 |
|---|---|
| 규칙 신규 등록 | rules.json에 규칙이 저장됨 |
| 규칙 수정 | 수정된 내용이 저장됨 |
| 규칙 삭제 | 삭제된 규칙이 목록에서 제거됨 |
| 규칙 비활성화 | 조건에 맞는 메일이 와도 자동 발송되지 않음 |
| 필수값 누락 | 저장되지 않고 사용자에게 안내됨 |
| 잘못된 이메일 주소 입력 | 저장 전 오류 안내 또는 저장 차단 |

---

## 18.2 자동 발송 테스트

| 테스트 항목 | 기대 결과 |
|---|---|
| 제목 조건 일치 메일 수신 | 자동 전달 메일 발송 |
| 제목 조건 불일치 메일 수신 | 아무 동작 없음 |
| 상단 메시지 삽입 | 전달 메일 본문 최상단에 저장된 메시지 삽입 |
| 수신자 설정 | 규칙에 저장된 수신자로 발송 |
| 참조자 설정 | 규칙에 저장된 참조자로 발송 |
| 숨은 참조자 설정 | 규칙에 저장된 숨은 참조자로 발송 |
| HTML 메일 처리 | HTML 형식 유지 후 상단 메시지 삽입 |
| 일반 텍스트 메일 처리 | 텍스트 형식으로 상단 메시지 삽입 |
| 첨부파일 포함 설정 | 첨부파일이 유지됨 |
| 첨부파일 제외 설정 | 전달 메일에서 첨부파일이 제거됨 |

---

## 19. 완료 기준

다음 조건을 모두 만족하면 개발 완료로 본다.

1. Outlook 실행 시 Add-in이 정상 로드된다.
2. 사용자가 규칙을 추가, 수정, 삭제할 수 있다.
3. 규칙 정보가 JSON 파일에 저장되고 다시 로드된다.
4. 신규 메일 수신 시 제목 조건을 정상적으로 비교한다.
5. 조건에 맞는 메일 수신 시 전달 메일이 자동 생성된다.
6. 규칙에 저장된 상단 메시지가 전달 메일 본문 최상단에 삽입된다.
7. 규칙에 저장된 수신자, 참조자, 숨은 참조자가 정상 반영된다.
8. 메일이 자동 발송된다.
9. 규칙 사용 여부에 따라 자동 발송이 제어된다.
10. 첨부파일 포함 여부 설정이 반영된다.
11. 로그 저장 기능은 포함하지 않는다.

---

## 20. 개발 시 주의사항

- Outlook Object Model 사용 시 COM 객체 해제에 주의한다.
- 자동 발송 기능은 조직 보안 정책에 따라 제한될 수 있다.
- HTML 본문에 사용자 메시지를 삽입할 때 반드시 HTML 인코딩을 수행한다.
- 여러 메일이 동시에 수신될 경우 이벤트가 연속 발생할 수 있으므로 중복 발송 가능성을 고려한다.
- 초기 버전에서는 첫 번째 매칭 규칙만 실행한다.
- 규칙 파일 경로가 없으면 자동으로 폴더와 파일을 생성한다.
- 규칙 파일 손상 시 Outlook Add-in 전체가 중단되지 않도록 예외 처리한다.
