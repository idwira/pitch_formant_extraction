from os import listdir
from xlwt import Workbook
import parselmouth

# Loading all sound files
SOUND_PATH = "C:\\tugas1\wav\\"
sound_files = sorted(listdir(SOUND_PATH))

# Preparing workbook
workbook = Workbook()
sheet = workbook.add_sheet("Sheet 1")
sheet.write(0, 0, "file_name")
sheet.write(0, 1, "mean_pitch")
sheet.write(0, 2, "mean_formant_1")
sheet.write(0, 3, "mean_formant_2")
sheet.write(0, 4, "mean_formant_3")
sheet.write(0, 5, "mean_formant_4")
sheet.write(0, 6, "mean_formant_5")

for index, sound_file in enumerate(sound_files):
    sheet.write(index + 1, 0, sound_file.split(".")[0])
    sound = parselmouth.Sound(SOUND_PATH + sound_file)

    pitch = sound.to_pitch()
    mean_pitch = parselmouth.praat.call(pitch, "Get mean", 0.0, 0.0, "Hertz")
    sheet.write(index + 1, 1, mean_pitch)

    formant = sound.to_formant_burg()
    for i in range(1, 6):
        mean_formant_i = parselmouth.praat.call(formant, "Get mean", float(i), 0.0, 0.0, "Hertz")
        sheet.write(index + 1, i + 1, mean_formant_i)

workbook.save("out_final.xls")
