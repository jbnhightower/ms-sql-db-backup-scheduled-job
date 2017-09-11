## userdefined varibles
$PathToOlaHallengren = 'C:\pathto\ola hallengren MaintenanceSolution.sql'
$BackupScriptLocation = "C:\pathto\ms-sql-db-backup-scheduled-job\sql backup assembly code output to transscript and variable.ps1"
$backuppath = 'C:\pathto\backup'
$sqlserver = 'somedomain\someinstance'
$username = 'somedomain\someuser'
$dailyTrigger = New-JobTrigger -Weekly -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -at "23:12"
#$dailyTrigger = New-JobTrigger -Daily -At "15:45"    # angiv tidspunkt for kørsel

## end userdefined varibles

$asbackuppath = '$backuppath = ' + "'$backuppath'"
$backupscript = Get-Content $BackupScriptLocation
$backupscript[1] = $asbackuppath
$backupscript | Out-File $BackupScriptLocation

function Get-WpfUserInput {
    
    [CmdletBinding()]
    Param
    (
        # Message for textbox
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $Message
    )
    
    Add-Type -AssemblyName PresentationFramework, System.Windows.Forms, WindowsFormsIntegration

    [xml][string]$XAML_ConnectDialog = @"
    <Window Name="Form_ConnectDialog"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="User indput" Height="250" Width="428" ResizeMode="NoResize" ShowInTaskbar="True" FocusManager.FocusedElement="{Binding ElementName=Txt_ConnectDialog_Input}">
        <Grid>
            <TextBlock FontSize="14" TextWrapping="Wrap" Text="$message" Margin="10,0,10,91"/>
            <Label FontSize="16" Content="Enter password:" HorizontalAlignment="Left" Height="35" VerticalAlignment="Top" Width="156" Margin="10,45,0,0"/>
            <Label FontSize="16" Content="Re-enter password:" HorizontalAlignment="Left" Height="35" VerticalAlignment="Top" Width="156" Margin="10,80,0,0"/>
            <Button Name="Btn_ConnectDialog_Connect"  HorizontalAlignment="center" Content="OK" Height="35" Width="100" Margin="50,88,26,20" IsDefault="True"/>
            <PasswordBox x:Name="Txt_ConnectDialog_Input" HorizontalAlignment="Left" Height="23" Margin="171,51,0,0" VerticalAlignment="Top" Width="208" />
            <PasswordBox x:Name="Txt_ConnectDialog_Input_control" HorizontalAlignment="Left" Height="23" Margin="171,87,0,0" VerticalAlignment="Top" Width="208" />
        </Grid>
    </Window>
"@
    [xml][string]$XAML_ConnectDialogPasswordmismatch = @"
    <Window Name="Form_ConnectDialog"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="User indput" Height="250" Width="428" ResizeMode="NoResize" ShowInTaskbar="True" FocusManager.FocusedElement="{Binding ElementName=Txt_ConnectDialog_Input}">
        <Grid>
            <TextBlock FontSize="14" TextWrapping="Wrap" Text="$message" Margin="10,0,10,91"/>
            <Label FontSize="16" Content="Enter password:" HorizontalAlignment="Left" Height="35" VerticalAlignment="Top" Width="156" Margin="10,45,0,0"/>
            <Label FontSize="16" Content="Re-enter password:" HorizontalAlignment="Left" Height="35" VerticalAlignment="Top" Width="156" Margin="10,80,0,0"/>
            <Label FontSize="16" Content="password mismatch" HorizontalAlignment="center" Height="35" VerticalAlignment="Top" Width="156" Margin="50,108,26,20"/>
            <Button Name="Btn_ConnectDialog_Connect"  HorizontalAlignment="center" Content="OK" Height="35" Width="100" Margin="50,108,26,20" IsDefault="True"/>
            <PasswordBox x:Name="Txt_ConnectDialog_Input" HorizontalAlignment="Left" Height="23" Margin="171,51,0,0" VerticalAlignment="Top" Width="208" />
            <PasswordBox x:Name="Txt_ConnectDialog_Input_control" HorizontalAlignment="Left" Height="23" Margin="171,87,0,0" VerticalAlignment="Top" Width="208" />
        </Grid>
    </Window>
"@
    $i = 0
    do {
        if ($i -gt 0) {
            $XML_Node_Reader_ConnectDialog = (New-Object System.Xml.XmlNodeReader $XAML_ConnectDialogPasswordmismatch)
        }
        else {
            $XML_Node_Reader_ConnectDialog = (New-Object System.Xml.XmlNodeReader $XAML_ConnectDialog)
        }
        #$XML_Node_Reader_ConnectDialog = (New-Object System.Xml.XmlNodeReader $XAML_ConnectDialog)
        $ConnectDialog = [Windows.Markup.XamlReader]::Load($XML_Node_Reader_ConnectDialog)
        $Btn_ConnectDialog_Connect = $ConnectDialog.FindName('Btn_ConnectDialog_Connect')
        $Txt_ConnectDialog_Input = $ConnectDialog.FindName('Txt_ConnectDialog_Input')
        $Txt_ConnectDialog_Input_control = $ConnectDialog.FindName('Txt_ConnectDialog_Input_control')
    
    
        $Btn_ConnectDialog_Connect.Add_Click( {
                $ConnectDialog.Close()
            })
    
        $ConnectDialog.Add_Closing( {[System.Windows.Forms.Application]::Exit()}) # {$form.Close()}
    
        # add keyboard indput
        [System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($ConnectDialog)
    
        # Running this without $appContext and ::Run would actually cause a really poor response.
        $ConnectDialog.Show()
    
        # This makes it pop up
        $ConnectDialog.Activate() | Out-Null
        #run the form ConnectDialog
        $appContext = New-Object System.Windows.Forms.ApplicationContext
        [System.Windows.Forms.Application]::Run($appContext)
    
        $UserName = "sa"
        $top = New-Object System.Management.Automation.PSCredential `
            -ArgumentList $UserName, $Txt_ConnectDialog_Input.SecurePassword
    
        $UserName = "sa"
        $bottom = New-Object System.Management.Automation.PSCredential `
            -ArgumentList $UserName, $Txt_ConnectDialog_Input_control.SecurePassword
        $i++
    }
    while (($top.GetNetworkCredential().Password) -cne ($bottom.GetNetworkCredential().Password))
    Write-Output $output
}
function Get-UserFromWellKnownSidType {
    [CmdletBinding()]

    [OutputType([psobject])]
    Param
    (
        # Welknown SID type from https://msdn.microsoft.com/en-us/library/system.security.principal.wellknownsidtype(v=vs.110).aspx
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [System.Security.Principal.WellKnownSidType]$WellKnownSidType = 'BuiltinUsersSid',

        # Domain SID, defaults to WellknownSIDType BuiltinDomainSid
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $domainSID = (New-Object System.Security.Principal.SecurityIdentifier([System.Security.Principal.WellKnownSidType]::BuiltinDomainSid, $null))

    )

    Begin {
        Push-Location
    }
    Process {

        $ID = [System.Security.Principal.WellKnownSidType]::$WellKnownSidType
        $SID = New-Object System.Security.Principal.SecurityIdentifier($ID, $domainSID)
        $objSID = New-Object System.Security.Principal.SecurityIdentifier($SID)
        $objUser = $objSID.Translate( [System.Security.Principal.NTAccount])
        [string]$NTAccount = $objUser.Value
        $props = @{ User = $NTAccount
            NTAccountSID = $SID
        }
        $UserFromWellKnownSidType = New-Object -TypeName psobject -Property $props
        Write-Output $UserFromWellKnownSidType

        Write-Output $SID
    }
    End {
        Pop-Location
    }
}
function Write-SqlcommandMessage {

    [CmdletBinding()]
    Param
    (
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $True)]
        [string]$SqlServer,

        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $False)]
        [string]$Database,

        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $True)]
        [string]$SqlStatement


    )
    begin {
        Push-Location
    }
    process {
        Write-Verbose "----------------------------------------------"
        Write-Verbose "Invoke-SqlSelect"
        Write-Verbose "SqlServer = $SqlServer"
        Write-Verbose "Database = $Database"
        $SqlStatementverbose = $SqlStatement -split "`n"
        $SqlStatementverbosefirst = $SqlStatementverbose[0]
        Write-Verbose "SqlStatement = $SqlStatementverbosefirst"
        $SqlStatementverbose = $SqlStatementverbose[1..$SqlStatementverbose.Length]
        foreach ($SqlStatementverboseline in  $SqlStatementverbose) {
            Write-Verbose "$SqlStatementverboseline"
        }
        Write-Verbose "----------------------------------------------"

        $ErrPrefence = $ErrorActionPreference
        $ErrorActionPreference = "Stop"

        $ConnectionString = "Server=$SqlServer;Database=$Database;Integrated Security=True"

        #assembly code
        # This is used because screen output from the .net obejct isn't captured by Start-Transcript.
        $source = @"
namespace Test
{
    using System;
    using System.Data;
    using System.Data.SqlClient;

    public class SomeProgram1
    {
        public static int Main(string[] args)
        {
            if (args.Length != 2)
            {
                Usage();
                return 1;
            }

            var conn = args[0];
            var sqlText = args[1];
            ShowSqlErrorsAndInfo(conn, sqlText);

            return 0;
        }

        private static void Usage()
        {
            Console.WriteLine("Usage: sqlServerConnectionString sqlCommand");
            Console.WriteLine("");
            Console.WriteLine("   example:  \"Data Source=.;Integrated Security=true\" \"DBCC CHECKDB\"");
        }

        public static void ShowSqlErrorsAndInfo(string connectionString, string query)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.StateChange += OnStateChange;
                connection.InfoMessage += OnInfoMessage;


                SqlCommand command = new SqlCommand(query, connection);
                try
                {
                    command.CommandTimeout = 0;
                    command.Connection.Open();
                    Console.WriteLine("Command execution starting.");
                    SqlDataReader dr = command.ExecuteReader();
                    if (dr.HasRows)
                    {
                        Console.WriteLine("Rows returned.");
                        while (dr.Read())
                        {
                            for (int idx = 0; idx < dr.FieldCount; idx++)
                            {
                                Console.Write("{0} ", dr[idx].ToString());
                            }

                            Console.WriteLine();
                        }
                    }

                    Console.WriteLine("Command execution complete.");
                }
                catch (SqlException ex)
                {
                    DisplaySqlErrors(ex);
                }
                finally
                {
                    command.Connection.Close();
                }
            }
        }

        private static void DisplaySqlErrors(SqlException exception)
        {
            foreach (SqlError err in exception.Errors)
            {
                Console.WriteLine("ERROR: {0}", err.Message);
            }
        }

        private static void OnInfoMessage(object sender, SqlInfoMessageEventArgs e)
        {
            foreach (SqlError info in e.Errors)
            {
                Console.WriteLine("INFO: {0}", info.Message);
            }
        }

        private static void OnStateChange(object sender, StateChangeEventArgs e)
        {
            Console.WriteLine("Connection state changed: {0} => {1}", e.OriginalState, e.CurrentState);
        }
    }
}
"@

        # load assembly above
        Add-Type -TypeDefinition $source -ReferencedAssemblies 'System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'

        #helper class to generate output"
        Add-Type -TypeDefinition @"
public class WriteTest {
public static void Write(string s) { System.Console.WriteLine(s);
    }
}
"@

        # create our string writer
        $writer = new-object io.stringwriter
        [Console]::SetOut($writer);

        # run assembly
        [Test.SomeProgram1]::ShowSqlErrorsAndInfo($ConnectionString, $SqlStatement) | out-null
        # finish work with the assembly
        $writer.Close()

        # restore standard output
        $standardOutput = new-object io.StreamWriter([Console]::OpenStandardOutput());
        $standardOutput.AutoFlush = $true;
        [Console]::SetOut($standardOutput);

        # check output
        $assemblyOutput = $writer.ToString()
        Write-Output $assemblyOutput
    }
    end {
        $ErrorActionPreference = $ErrPrefence
        Pop-Location
    }
}
function Invoke-SqlSelect {
    
    [CmdletBinding()]
    Param
    (
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $True)]
        [string]$SqlServer,
    
        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $False)]
        [string]$Database,
    
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $True)]
        [string]$SqlStatement,
    
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipelineByPropertyName = $True, Mandatory = $false)]
        [string]$timeout
    )
    begin {
        Push-Location
    }
    process {
        Write-Verbose "----------------------------------------------"
        Write-Verbose "Invoke-SqlSelect"
        Write-Verbose "SqlServer = $SqlServer"
        Write-Verbose "Database = $Database"
        Write-Verbose "timeout = $timeout"
        $SqlStatementverbose = $SqlStatement -split "`n"
        $SqlStatementverbosefirst = $SqlStatementverbose[0]
        Write-Verbose "SqlStatement = $SqlStatementverbosefirst"
        $SqlStatementverbose = $SqlStatementverbose[1..$SqlStatementverbose.Length]
        foreach ($SqlStatementverboseline in  $SqlStatementverbose) {
            Write-Verbose "$SqlStatementverboseline"
        }
        Write-Verbose "----------------------------------------------"
    
        $ErrorActionPreference = "Stop"
    
        $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $sqlConnection.ConnectionString = "Server=$SqlServer;Database=$Database;Integrated Security=True"
    
        $sqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $sqlCmd.CommandText = $SqlStatement
        $sqlCmd.Connection = $sqlConnection
        $sqlCmd.CommandTimeout = $timeout
    
    
        $sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $sqlAdapter.SelectCommand = $sqlCmd
        $dataTable = New-Object System.Data.DataTable
        try {
            $sqlAdapter.SelectCommand.Connection.Open()
            $sqlOutput = $sqlAdapter.Fill($dataTable)
            $sqlAdapter.SelectCommand.Connection.Close()
            $sqlAdapter.Dispose()
            Write-Output -Verbose $sqlOutput
        }
        catch {
            Write-Warning "Error executing SQL on database [$Database] on server [$SqlServer]." #Statement: `r`n$SqlStatement"
            Write-Verbose $error[0]
            return $null
        }
    
        if ($dataTable) { return , $dataTable } else { return $null }
    }
    end {
        Pop-Location
    }
}

$windowsuserpassword = Get-WpfUserInput -Message 'Enter password for windows user'

$usercred = New-Object System.Management.Automation.PSCredential `
    -ArgumentList $UserName, $windowsuserpassword

#$usercred = Get-Credential -UserName "$env:COMPUTERNAME\$username" -Message 'tast kode til visma brugeren der bliver oprettet til backup'     

$userinput = Get-WpfUserInput -Message 'Enter Password for sa'

Invoke-Sqlcmd -InputFile $PathToOlaHallengren -ServerInstance $sqlserver -Username 'sa' -Password $($BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($userinput); [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR))

$administrator = Get-UserFromWellKnownSidType -WellKnownSidType BuiltinAdministratorsSid 

$admingroup = ($administrator.user).TrimStart('BUILTIN\\')

#[securestring]$password = Read-Host -AsSecureString

try {
    New-LocalUser -AccountNeverExpires -Name visma -Description 'Bruges af visma til sql backup' -PasswordNeverExpires -Password ($usercred.Password)
}
catch [System.Management.Automation.CommandNotFoundException] {
    Write-Verbose "new-localuser is not supported running NET USER instead"
    
    [string]$string = $($BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($usercred.Password); [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)) 

    NET USER $username "$string" /ADD
}


try {
    Add-LocalGroupMember -Group "$admingroup" -Member "$usercred.UserName" 
}
catch [System.Management.Automation.CommandNotFoundException] {
    Write-Verbose "Add-LocalGroupMember is not supported running NET LOCALGROUP instead" 

    NET LOCALGROUP "$admingroup" "$username" /ADD
}

<#
$group = [ADSI]"WinNT://gamer/visma,group"
$group
#>

$option = New-ScheduledJobOption -StartIfOnBattery

Register-ScheduledJob -Name databasebackup -FilePath $BackupScriptLocation -Trigger $dailyTrigger -ScheduledJobOption $option -Credential $usercred

<#
Register-ScheduledJob 
(Get-ScheduledJob -Name databasebackup).Run()
#>

$domainuser = $usercred.UserName

$AddWindowsUserquery = @"
CREATE LOGIN [$domainuser] FROM WINDOWS;  
"@

$AddSqlUserAddSysadmingroupQuery2012Forward = @"
ALTER server ROLE sysadmin ADD MEMBER [$domainuser]
"@

$AddSqlUserAddSysadmingroupQuery2008only = @"
EXEC master..sp_addsrvrolemember @loginame = N'$domainuser', @rolename = N'sysadmin'
"@

$sqlquerysqlserver2012Forward = [ordered]@{AddWindowsUserquery = $AddWindowsUserquery
    AddSqlUserAddSysadmingroupQuery2012Forward = $AddSqlUserAddSysadmingroupQuery2012Forward
}

$sqlquerysqlserver2008Only = [ordered]@{AddWindowsUserquery = $AddWindowsUserquery
    AddSqlUserAddSysadmingroupQuery2008only = $AddSqlUserAddSysadmingroupQuery2008only
}

$querysqlversion = @"
SELECT  
SERVERPROPERTY('ProductVersion') AS ProductVersion;
"@

$res = Invoke-SqlSelect -SqlServer $sqlserver -Database 'master' -SqlStatement $querysqlversion 

[int]$SqlVersionMajor = ($res.ProductVersion).Substring('0', '2')

if ($SqlVersionMajor -ge 11) {
    foreach ($query in $sqlquerysqlserver2012Forward.Values) {
        write-SqlcommandMessage -SqlServer $SqlServer -Database 'master' -SqlStatement $query -Verbose
    }
}
elseif ($SqlVersionMajor -lt 11) {
    foreach ($query in $sqlquerysqlserver2008Only.Values) {
        write-SqlcommandMessage -SqlServer $SqlServer -Database 'master' -SqlStatement $query -Verbose
    }
}

$SqlSysadminQuery = @"
SELECT sys.server_role_members.role_principal_id, role.name AS RoleName,   
sys.server_role_members.member_principal_id, member.name AS MemberName  
FROM sys.server_role_members  
JOIN sys.server_principals AS role  
ON sys.server_role_members.role_principal_id = role.principal_id  
JOIN sys.server_principals AS member  
ON sys.server_role_members.member_principal_id = member.principal_id;  
"@

$ResSqlSysadmin = Invoke-SqlSelect -SqlServer $sqlserver -Database 'master' -SqlStatement $SqlSysadminQuery 

if ($ResSqlSysadmin.MemberName -ccontains $domainuser) {
    Write-Verbose "{$domainuser} is added to sysadmins"
}
else {
    Write-warning "{$domainuser} is NOT added to sysadmins"
}

<#
[int]$i = 0
foreach ($query in $sqlquerysqlserver2012Forward.Values) {
    $i++
    $res = invoke-sqlselect -SqlServer $SqlServer -Database 'master' -SqlStatement $query -Verbose
    New-Variable -Name "var$i" -Value $result
    write-verbose "result for 'var'$i var$i"
}

$AddSqlUserQuery2008backward = @"
USE [master]
CREATE LOGIN&nbsp;[$domainuser] WITH PASSWORD=N'test', DEFAULT_DATABASE=[master], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF
"@

$SqlShowAccessForUser = @" 
EXEC sp_addrolemember 'sysadmin', '[$domainuser]'
SELECT * FROM sys.fn_builtin_permissions('SERVER') ORDER BY permission_name;
EXEC sp_addrolemember 'sysadmin', '[$domainuser]'

"@

$SqlSysadminQuery = @"
SELECT * FROM sys.fn_builtin_permissions('SERVER') ORDER BY permission_name;
"@
#>