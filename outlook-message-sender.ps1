$fromaddress = "info@domain.com";  #needs to be configured within Your Outlook app
function Invoke-SetProperty {
    param(
        [__ComObject] $Object,
        [String] $Property,
        $Value
    )
    [Void] $Object.GetType().InvokeMember($Property,"SetProperty",$NULL,$Object,$Value)
}
$Config = Get-Content -Path "C:\_DEV\config.txt"|ForEach-Object -Process {
    $Config = $_;
    $ConfigArray = $Config.Split("_");
    $Name=$ConfigArray['0'].Trim();
    $Value=$ConfigArray['1'].Trim();
    if($Name -eq "DateTime"){$DateTime = $Value};
    if($Name -eq "TimeZone"){$TimeZone = $Value};
    if($Name -eq "EstTime"){$EstTime = $Value};
    if($Name -eq "EnvAffected"){$EnvAffected = $Value};
    if($Name -eq "Services"){$Services = $Value};
    $Subject = "Maintenance Window $DateTime "+ $TimeZone;
}
$htmltext = "<!DOCTYPE html><HTML><HEAD><STYLE>.headings1{font-family: Verdana; font-size: 22px; font-stretch: expanded; letter-spacing: 0px; word-spacing: 0px; color: #000000; margin:auto; text-align:left; font-weight: 400; text-decoration: none; font-style: normal; padding-bottom:30px;}.headings2{font-family: Verdana; font-size: 16px; letter-spacing: 0px; word-spacing: 1px; color: #000000; font-weight: 100; text-decoration: none; font-style: normal;}.headings3{font-family: Verdana; font-size: 12px; letter-spacing: 0px; word-spacing: 2px; color: #cccccc; font-weight: 100; text-decoration: none; font-style: normal;}.value{color:rgb(0, 0, 0); font-family: Verdana; font-size: 12px; letter-spacing: 0px; word-spacing: 0px; font-weight: 150; text-decoration: none; font-style: normal;}.notification{font-family: Verdana; font-size: 12px; letter-spacing: 0px; word-spacing: 0px; color: #000000; font-weight: 100; text-decoration: none; font-style: normal;}table{border: 1px solid #CCCCCC; margin:auto; margin-top: 40px; width: 600px; height:fit-content; border-spacing: 20px; padding-left: 20px; padding-right: 20px; padding-bottom: 20px; padding-top: 40px;}tr, th, td{height:fit-content;}</style></HEAD><BODY> <CENTER> <TABLE class='table'> <tr><td><H1 class='headings1'>Your Valued Company Scheduled Maintenance</H1><p class='headings2'>Upcoming scheduled maintenance notice</p><p class='notification'>This notice is to inform you that we will be performing scheduled maintenance that will affect the availability of <span style='color:rgb(0, 0, 0);font-weight:150px'>Product1 & Product 2 services</span>.</p></td></tr><tr><td><p class='notification'>Impact:<br/>During the execution of the scheduled activity, customer may experience the inability to establish call to our agents. No customer intervention is required.</p></td></tr><tr><td><h3 class='headings3'>Start time</h3><p class='value'>DateTime</p></td></tr><tr><td><h3 class='headings3'>Estimated duration</h3><p class='value'>EstTime</p></td></tr><tr><td><h3 class='headings3'>Environments affected</h3><p class='value'>EnvAffected</p></td></tr><tr><td><h3 class='headings3'>Components affected</h3><p class='value'>Services</p><br/></td></tr></TABLE> </CENTER></BODY></HTML>";
$resultingtext1 = $htmltext.replace('DateTime', $DateTime);
$resultingtext2 = $resultingtext1.replace('TimeZone', $TimeZone);
$resultingtext3 = $resultingtext2.replace('EstTime', $EstTime);
$resultingtext4 = $resultingtext3.replace('EnvAffected', $EnvAffected);
$resultingtext5 = $resultingtext4.replace('Services', $Services);
Get-Content -Path "C:\_DEV\addresses.txt"|ForEach-Object -Process {
    $Adresa = $_;
    $CharArray = $Adresa.Split(",");$mail=$CharArray['0'].Trim();
    #$name=$CharArray['1'].Trim();
    $outlook = new-object -comobject outlook.application;
    $email = $outlook.CreateItem(0);
    $account = $email.Session.Accounts.Item($fromaddress);
    Invoke-SetProperty -Object $email -Property "SendUsingAccount" -Value $account;
    $email.To = $mail;
    $email.Subject = $Subject;
    $priloha=$CharArray['2'].Trim();
    if($priloha -eq "0"){
        #-- no attachment added"
    }else {
        $email.Attachments.Add("C:\_DEV\addresses.txt");
    };
    $email.HTMLBody = $resultingtext5;
    $email.Send();
}