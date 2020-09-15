
Import-Module "C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"

$client = [Microsoft.Lync.Model.LyncClient]::GetClient()
$self = $client.Self
$conversationMgr = $client.ConversationManager

# Get a list of all open convesations