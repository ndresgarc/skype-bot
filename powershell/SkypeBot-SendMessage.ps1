
# USE
#  .\SkypeBot-SendMessage.ps1 -contact "email@domain.com" -text "Hello"

param([string] $contact = "Name.Surname@bshg.com", 
      [string] $text = "Hello")
	  
Import-Module "C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.dll"

$client = [Microsoft.Lync.Model.LyncClient]::GetClient()
$conversation = $client.ConversationManager.AddConversation()
$person = $client.ContactManager.GetContactByUri($contact)
$conversation.AddParticipant($person)

$PlainText = 0
$IMType = 1
$dictionary = New-Object "System.Collections.Generic.Dictionary[Microsoft.Lync.Model.Conversation.InstantMessageContentType,String]"
$dictionary.Add($PlainText, $text)

$message = $conversation.Modalities[$IMType]

$null = $message.BeginSendMessage($dictionary, $null, $dictionary)