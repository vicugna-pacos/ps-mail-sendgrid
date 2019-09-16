<#
    SendGrid Web API v3�Ń��[���𑗂�B

    API�h�L�������g�Fhttps://sendgrid.api-docs.io/v3.0/mail-send/v3-mail-send
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

    # �w�b�_�[�̓��e��API�L�[��content-type�łقڌŒ�
    $headers = @{
        "authorization" = "Bearer ${API_KEY}"
        "content-type" = "application/json"
    }

    $body = @{}

    # ����
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

    # �����Ɩ{��
    $content = @{
        "type" = "text/plain"
        "value" = $Message
    }
    $body.Add("subject", $Subject)
    $body.Add("content", @($content))

    # From
    $body.Add("from", @{"email" = $From})

    # �Y�t�t�@�C��
    $b_attr = AddAttachments($Attachments)

    if ($b_attr -ne $null) {
        $body.Add("attachments", $b_attr)
    }

    # �I�u�W�F�N�g��Json�֕ϊ�����B���̍ہADepth���w�肵�āA�[���K�w���ϊ������悤�ɂ���B
    $bodyJson = $body | ConvertTo-Json -Depth 20

    # Json��UTF-8�֕ϊ�����B�ϊ�������POST����ƁA���{�ꂪ������������B
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($bodyJson)

    $res = Invoke-RestMethod -Uri $API_URL -Method Post -Headers $headers -Body $bodyBytes
}

<#
    ���[�����M��� personalizations �ɒǉ�����B
    to, cc, bcc��������kbn�Ŏw�肷��B
#>
function AddAddress([string[]]$addressList) {
    
    $email = @()

    foreach ($address in $addressList) {
        $email += @{"email" = $address}
    }

    return $email
}

<#
    �Y�t�t�@�C����ǉ�����
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
