from gtts import gTTS
from docx import Document
from pydub import AudioSegment
import io

AudioSegment.ffmpeg = r"C:\ffmpeg\ffmpeg-7.1-essentials_build\bin\ffmpeg.exe"  # Replace with your path

def extract_columns_and_generate_audio(file_path, output_file="vocab_audio.mp3"):
    # Load the Word document
    doc = Document(file_path)

    # Extract rows from the table
    vocab_list = []
    for table in doc.tables:
        for row in table.rows[1:]:  # Skip the header row
            if len(row.cells) >= 3:
                word_type = row.cells[0].text.strip()  # First column: Type of word
                german_word = row.cells[1].text.strip()  # Second column: German word
                english_word = row.cells[2].text.strip()  # Third column: English word
                if german_word and english_word and word_type:
                    vocab_list.append((word_type, german_word, english_word))

    # Generate individual speech files directly in memory
    audio_segments = []
    silence = AudioSegment.silent(duration=2000)  # 2 seconds of silence
    short_silence = AudioSegment.silent(duration=700)  # 700 milliseconds of silence

    for word_type, german, english in vocab_list:
        # print(f"adding {word_type} {german} {english}")
        
        # Generate speech for word type (in memory)
        type_fp = io.BytesIO()
        tts_type = gTTS(word_type, lang="de")  # Word type in German
        tts_type.write_to_fp(type_fp)
        type_fp.seek(0)  # Reset pointer to the beginning of the stream
        type_audio = AudioSegment.from_file(type_fp, format="mp3")

        # German word audio (in memory)
        german_fp = io.BytesIO()
        tts_german = gTTS(german, lang="de")
        tts_german.write_to_fp(german_fp)
        german_fp.seek(0)  # Reset pointer to the beginning of the stream
        german_audio = AudioSegment.from_file(german_fp, format="mp3")

        # English word audio (in memory)
        english_fp = io.BytesIO()
        tts_english = gTTS(english, lang="en-us")
        tts_english.write_to_fp(english_fp)
        english_fp.seek(0)  # Reset pointer to the beginning of the stream
        english_audio = AudioSegment.from_file(english_fp, format="mp3")

        # Combine Word Type, German word, silence, and English word
        audio_segments.append(type_audio + short_silence + german_audio + silence + english_audio + silence)

    # Concatenate all audio segments
    full_audio = sum(audio_segments)

    # Export the final audio
    full_audio.export(output_file, format="mp3")
    print(f"Audio saved to {output_file}")

# Usage
extract_columns_and_generate_audio("VocabGerman.docx")
