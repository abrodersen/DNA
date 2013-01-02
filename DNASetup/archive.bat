 7z a -t7z dna.7z DNASetup.msi setup.exe
 copy /b 7zsd_All.sfx + config.txt + dna.7z dna_setup.exe /Y
 del dna.7z DNASetup.msi setup.exe