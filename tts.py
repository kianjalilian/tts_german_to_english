from gtts import gTTS
from docx import Document
from pydub import AudioSegment
AudioSegment.ffmpeg = r"C:\ffmpeg\ffmpeg-7.1-essentials_build\bin\ffmpeg.exe"  # Replace with your path

def extract_columns_and_generate_audio(file_path, output_file="vocab_audio.mp3"):
    # Load the Word document
    doc = Document(file_path)

    # Extract rows from the table
    vocab_list = []
    for table in doc.tables:
        for row in table.rows[1:]:  # Skip the header row
            if len(row.cells) >= 3:
                german_word = row.cells[1].text.strip()  # Second column
                english_word = row.cells[2].text.strip()  # Third column
                if german_word and english_word:
                    vocab_list.append((german_word, english_word))

    # Generate individual speech files
    audio_segments = []
    silence = AudioSegment.silent(duration=2000)  # 2 seconds of silence

    for german, english in vocab_list:
        print("adding ", german, english)
        # German audio
        tts_german = gTTS(german, lang="de")
        german_audio_path = "german_temp.mp3"
        tts_german.save(german_audio_path)
        german_audio = AudioSegment.from_file(german_audio_path)

        # English audio
        tts_english = gTTS(english, lang="en")
        english_audio_path = "english_temp.mp3"
        tts_english.save(english_audio_path)
        english_audio = AudioSegment.from_file(english_audio_path)

        # Combine German, silence, and English
        audio_segments.append(german_audio + silence + english_audio + silence)

    # Concatenate all audio segments
    full_audio = sum(audio_segments)

    # Export the final audio
    full_audio.export(output_file, format="mp3")
    print(f"Audio saved to {output_file}")

# Usage
extract_columns_and_generate_audio("VocabGerman.docx")
