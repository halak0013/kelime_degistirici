#! /bin/bash
echo "gerekli paketler yükleniyor..."
sudo apt install python3-openpyxl python3-docx python3-odf

chmod +x kelime_degistirici.py

echo "dosyalar taşınıyor..."
mkdir -p ~/.local/share/applications/
mkdir -p ~/.local/bin/

cp kelime_degistirici.py ~/.local/bin/

cp kelime_degistirici.desktop ~/.local/share/applications/
