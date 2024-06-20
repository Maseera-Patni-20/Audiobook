<!DOCTYPE html>
<html>
<body>

<h1>DOCX to Speech Converter</h1>

<p>This Python script reads text from a <code>.docx</code> file and converts it into speech, saving the audio output as an <code>mp3</code> file.</p>

<h2>Requirements</h2>
<ul>
    <li>Python 3</li>
    <li><a href="https://pypi.org/project/python-docx/">python-docx</a></li>
    <li><a href="https://pypi.org/project/pyttsx3/">pyttsx3</a></li>
</ul>

<h2>Installation</h2>
<pre><code>pip install python-docx pyttsx3</code></pre>

<h2>Usage</h2>
<pre><code>import docx
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
</code></pre>
