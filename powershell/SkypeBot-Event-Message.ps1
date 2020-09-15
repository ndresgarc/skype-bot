
Import-Module "C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"

$client = [Microsoft.Lync.Model.LyncClient]::GetClient()
$self = $client.Self
$conversationMgr = $client.ConversationManager

# Unregister old events subscribers
Get-EventSubscriber|Unregister-Event


# Register events for existing conversations.

$i = 0

for ($i=0; $i -lt $conversationMgr.Conversations.Count; $i++) {
	for ($j=0; $j -lt $conversationMgr.Conversations[$i].Participants.Count; $j++) {
	
			if ($conversationMgr.Conversations[$i].Participants[$j].IsSelf) {
			                   Write-Host "DEBUG: Skipping register for ANDRES"
            
			} else {
			
			Register-ObjectEvent -InputObject $conversationMgr.Conversations[$i].Participants[$j].Modalities[1] -EventName "InstantMessageReceived" `
                    -SourceIdentifier "new im $i - $j" `
                    -action {
                                                        $message = $EventArgs.Text
                               Write-Host "DEBUG: New incoming IM - $message"
                            }
			}
	
	
	}
}



# If this exits, no more events.
while (1) { }