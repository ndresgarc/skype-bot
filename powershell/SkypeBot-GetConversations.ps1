
Import-Module "C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"

$client = [Microsoft.Lync.Model.LyncClient]::GetClient()
$self = $client.Self
$conversationManager = $client.ConversationManager

$conversationObjectArray = @()

foreach ($conversation in $conversationManager.Conversations) { 
	$conversationObject = New-Object -TypeName psobject
	foreach ($participant in $conversation.Participants) {
		
		if (!($participant.IsSelf)) {
			foreach ($property in $participant.Properties) {
				if ($property.key -eq "Name") {
					$property.Value
					$conversationObject | Add-Member -MemberType NoteProperty -Name User -Value $property.Value
				} 
			}
		
		}
	
	}
	$conversationObjectArray += $conversationObject
}

$conversationObjectArray | ConvertTo-Json