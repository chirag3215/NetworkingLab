# NetworkingLab

This product was deisgned to expedite the data collection process for the Networking Lab at UC San Diego.

The task was to transcribe the images of participants networking messages to a google spreadsheet
to diagnose if the messages would be a deciding factor in response rate.

The images would follow a naming convention along the lines of "R_0IMURLff7tWgOFb_IMG_1313.png" where 
"0IMURLff7tWgOFb" is the participant ID that we want to match to the transcribed image. So we would first 
run **renameFilesInFolderWithParticipantID**.js to extract the participant id from the filenames and then we 
run **transcribeFilesInOrder**.js to match the participant id to the file which we are going to transcribe by 
using **Google Cloud Vision API** to scan the images of all text.

Because the text reader takes in all text and the screenshots don't follow a uniform format, we need to have extra 
supporting scripts to clean the text and extract the netowkring message. What was most effective was **extractCoreMessages**.js
for where the message used a common greeting like "dear" or "hello" to mark the beginning of the message and to look 
for words like "sincerely", or "thank you" to to mark the end of the message (while also taking the next word as 
it is likely to be the name of the message sender).

Additionally we want to quickly identify if the particiapants followed the script they were given which were either:
[
Dear [Name],
How are you? [Share your passion]. I am very inspired by your role as [Insert career here]. I would be honored to learn more
about your career journey. Are you available for a 15-minute informational interview on [insert a day that's 3-4 days away] 
at [give a specific time] or [second option date] between [give a time range]? I look forward to hearing from you!
Sincerely,
[Your Name]

Dear [Name],
I am very inspired by your role as [Insert career here]. I would be honored to learn more about your career journey. Are you available for a 15-minute informational interview on [day] at [time] or [day] between [time and time]? I look forward to hearing from you!
Sincerely,
[Your Name]
]

Both messages used the word honored and inpsired which seem to be irregular from normal vernacular so we simply
marked all messages that contained either of the words as following the script using **checkMessagesForKeywords**.js


