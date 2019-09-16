<#
    SendGrid Web API v3でメールを送る。

    APIドキュメント：https://sendgrid.api-docs.io/v3.0/mail-send/v3-mail-send
#>

$API_URL = "https://api.sendgrid.com/v3/mail/send"
$API_KEY = "<<API KEY>>"

function SendMail(
        [string]$From = "from@example.com",
        [string[]]$To = $null,
        [string[]]$Cc = $null,
        [string[]]$Bcc = $null,
        [string]$Subject = $null,
        [string]$Message = $null,
        [string[]]$Attachments = $null) {

    # ヘッダーの内容はAPIキーとcontent-typeでほぼ固定
    $headers = @{
        "authorization" = "Bearer ${API_KEY}"
        "content-type" = "application/json"
    }

    $body = @{}

    # 宛先
    $b_per = @{}

    if ($To -ne $null) {
       [object[]]$wk = AddAddress($To)
       $b_per.Add("to", $wk)
    }

    if ($Cc -ne $null) {
       [object[]]$wk = AddAddress($Cc)
       $b_per.Add("cc", $wk)
    }

    if ($Bcc -ne $null) {
       [object[]]$wk = AddAddress($Bcc)
       $b_per.Add("bcc", $wk)
    }

    $body.Add("personalizations", @($b_per))

    # 件名と本文
    $content = @{
        "type" = "text/plain"
        "value" = $Message
    }
    $body.Add("subject", $Subject)
    $body.Add("content", @($content))

    # From
    $body.Add("from", @{"email" = $From})

    # 添付ファイル
    $b_attr = AddAttachments($Attachments)

    if ($b_attr -ne $null) {
        $body.Add("attachments", $b_attr)
    }

    # オブジェクトをJsonへ変換する。その際、Depthを指定して、深い階層も変換されるようにする。
    $bodyJson = $body | ConvertTo-Json -Depth 20

    # JsonをUTF-8へ変換する。変換せずにPOSTすると、日本語が文字化けする。
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyJson)

    $res = Invoke-RestMethod -Uri $API_URL -Method Post -Headers $headers -Body $bodyBytes
}

<#
    メール送信先を personalizations に追加する。
    to, cc, bccを引数のkbnで指定する。
#>
function AddAddress([string[]]$addressList) {
    
    $email = @()

    foreach ($address in $addressList) {
        $email += @{"email" = $address}
    }

    return $email
}

<#
    添付ファイルを追加する
#>
function AddAttachments([string[]]$filenames) {
    if ($filenames -eq $null) {
        return $null
    }

    $attachments = @()

    foreach ($filename in $filenames) {
        $base64 = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($filename))
        $file = Get-ChildItem $filename

        $attachments += @{
            "content" = $base64
            "filename" = $file.Name
        }
    }

    return $attachments
}
