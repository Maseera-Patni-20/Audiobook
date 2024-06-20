import docx
import pyttsx3

# Open the DOCX file
doc = docx.Document('SPEECH WCD.docx')

# Initialize the text-to-speech engine
speaker = pyttsx3.init()

# Read all the paragraphs in the DOCX file
full_text = []
for para in doc.paragraphs:
    full_text.append(para.text)

# Concatenate all the text
text = '\n'.join(full_text)

# Have the speaker say the text
speaker.say(text)
speaker.runAndWait()

# Save the extracted text to an audio file
speaker.save_to_file(text, 'audio.mp3')
speaker.runAndWait()

# Stop the speaker
speaker.stop()
