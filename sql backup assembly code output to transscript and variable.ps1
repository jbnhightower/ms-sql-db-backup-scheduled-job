$date = get-date -Format yyyyMMddhhmm
$backuppath = 'C:\visma\backup'
Start-Transcript -Path "$backuppath\transcript$date.txt"

#userdefined varibles

$instance = 'winkompas2012'
[int]$antalversionerafdatabasen = 3
[int]$antalversioneraftrancescripts = 100 

#end userdefined varibles
$backuppath = $backuppath.TrimEnd('\\')
$backuppathextension = "$env:COMPUTERNAME`$$instance"

# connectionstring and query
$SqlServer = "$env:COMPUTERNAME\$instance"
$Database = "master" 
$connectionstring = "Server=$SqlServer;Database=$Database;Integrated Security=True"
$query = @"
EXECUTE dbo.DatabaseBackup
@Databases = 'USER_DATABASES',
@Directory = '$backuppath',
@BackupType = 'FULL',
@Verify = 'Y',
@CheckSum = 'Y'
"@

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
        Write-Verbose "Write-SqlcommandMessage"
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

write-SqlcommandMessage -SqlServer $SqlServer -Database $Database -SqlStatement $query -Verbose

#transscripts oprydning. der bliver lavet et transscript hver gang. For ikke at sande disken til med log filer, fjerner jeg de gamle når vi når det definerede antal.
$match = Get-ChildItem $backuppath -Filter '*transcript*.txt'| Sort-Object -Property lastwritetime -Descending 
if ($match.Count -ge $antalversioneraftrancescripts) {
    $o = $antalversioneraftrancescripts
    $sel = For (; $o -le $match.Count - 1; $o++ ) {
        $o
    }
    foreach ($sin in $sel) {
        ($match[$sin].fullname)  | Remove-Item -force
    }
}

[string[]]$databasefolders = Get-ChildItem "$backuppath\$backuppathextension" -Recurse -Filter '*.bak' | Select-Object -ExpandProperty PSParentPath | Get-Unique
foreach ($folder in $databasefolders) {
    $match = Get-ChildItem $folder -Filter '*.bak' | Sort-Object -Property lastwritetime -Descending 
    if ($match.Count -ge $antalversionerafdatabasen) {
        $o = $antalversionerafdatabasen
        $sel = For (; $o -le $match.Count - 1; $o++ ) {
            $o
        }
        foreach ($sin in $sel) {
            ($match[$sin].fullname)  | Remove-Item -for
        }
    }
}



Stop-Transcript

<#
$TypeName =  'System.Data'
$TypeName = 'System.Data.SqlClient'
[System.Reflection.Assembly]::LoadWithPartialName($TypeName).Location;
[System.Reflection.Assembly]::LoadWithPartialName($TypeName).fullname

Needs ola hallengren stored procedure to work
sysadmin is needed for the sql server



Import-Module sqlps -disablenamechecking
Invoke-Sqlcmd -InputFile "C:\Visma\ola hallengren MaintenanceSolution.sql" -ServerInstance $sqlserver -Username sa -Password password
Invoke-Sqlcmd -InputFile "C:\Visma\ola hallengren MaintenanceSolution.sql" -ServerInstance $sqlserver -Database 'master'
Set-ExecutionPolicy RemoteSigned
Set-Location c:

$res = Invoke-Sqlcmd -Query $query -ServerInstance .\'WINKOMPAS2012' -Username sa -Password passsword
 

$res | Where-Object -FilterScript {$_.name -like '*backup*'}

$query = @"
SELECT name, 
       type
  FROM dbo.sysobjects
 WHERE (type = 'P')
"@


$query = @"
SELECT name, 
       type
  FROM dbo.sysobjects
 WHERE (type = 'P')
"@

$res = Invoke-Sqlcmd -Query $query -ServerInstance .\'WINKOMPAS2012' -Username sa -Password password

$dbselect = (Get-ChildItem SQLSERVER:\SQL\$env:COMPUTERNAME\winkompas2012\Databases).name


#>


