# Attachments downloader for Microsoft Exchange as Windows service 
 I  would like to introduce one my small project. May be at once it will be usefull for some body. At one day I have got task to find solution for
 downloading attachments in certain mailbox and then put these file to folder in fileserver. Name folder has to depend on processing date. Processing
 these files from atach greatly affects the business process in company. Therefore this solution has to be reliable and works without any intervation 
 by admins. According to requirements I decide to write application as Windows Service that will be use Exchange Web Services. 
 This project was created in Microsoft Visual Studio 2010 and was tested in Microsoft Exchange 2010 SP1 enviroment.
 After build project you have to install AttachmentsDownloader.exe as service. For this purpose use  
 AttachmentsDownloader\bin\Release\installService.cmd but before edit it depends on target platform.
 	Also the Exchange Web Services (EWS) Managed API 2.0 must be installed on your pc. In project folder AttachmentsDownloader\bin\Release\ 
 located two dll that used in developing and testing enviroment.
	 In order to configure custom applicantion settings  change  section  <appSetting> in config AttachmentsDownloader.exe.config


<appSettings>
    <add key="MailServerUrl" value="https://mail.contoso.com/EWS/Exchange.asmx"/>
    <add key="MailBox" value="viktor@contoso.com"/>
    <add key="User" value="viktor"/>
    <add key="Password" value="Pa$$word"/>
    <add key="Domain" value="contoso.com"/>

    <!-- Destination root folder to save files in UNC format -->
    <add key="FolderRoot" value="\\fileserver.contoso.com\d$\tmp\vrf\"/>
 
    <!-- How much emails retrieve to proccess -->
    <add key="ItemsProccesing" value="50" />

    <!--Delay in miliseconds between scan Inbox-->
    <add key="ScanDelay" value="15000" />

    <!-- false - folder format \YYYY\MMYY\DDMMYYYY -->
    <!-- true  - folder format \YYYY\MMYY\DDMMYYYY\SenderEmailDomain -->
    <add key="FolderDestinationNameWithDomain" value="true" />

    <!-- false - leave original file name -->
    <!-- true  - FileName + ddMMyyyy_HHmmss+FileExtesion -->
    <add key="DestinationFileFormatLong" value="true" />

    <add key="ProcessingFileToLog" value="false" />
    <add key="ClientSettingsProvider.ServiceUri" value="" />
  </appSettings>
 

   