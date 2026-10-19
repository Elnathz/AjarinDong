# AQG mT5-base Fase 2: training dan local quiz pipeline

Bundle ini berisi notebook retraining Kaggle dan pipeline lokal untuk ekstraksi dokumen serta pembentukan kuis.

## Isi bundle

- `mt5_aqg_phase2_kaggle.ipynb`: revisi notebook yang diunggah. Tetap menggunakan `google/mt5-base`, DDP melalui `accelerate.notebook_launcher`, evaluasi `eval_strategy="steps"`, ROUGE-1/2/L, BLEU-4, serta answer-sensitivity audit.
- `aqg_local_pipeline.py`: satu script mandiri untuk PDF/PPTX/DOCX, normalisasi matematika, ekstraksi calon jawaban, inferensi batch mT5 lokal, filtering, distractor, dan output JSON.
- `test_aqg_local_pipeline.py`: unit test untuk fungsi-fungsi kritis.
- `aqg_demo_output.json`: contoh output demo tanpa bobot model.
- `requirements_local.txt`: dependensi parser dan inference lokal.

## A. Retraining Kaggle

1. Buka `mt5_aqg_phase2_kaggle.ipynb` di Kaggle dengan Internet aktif dan accelerator **2× Tesla T4**.
2. Jalankan cell dari atas secara berurutan. Output lama sudah dibersihkan agar tidak dianggap sebagai hasil dari kode Fase 2.
3. Notebook memuat `rifkiaputri/idk-mrc` dan `rajpurkar/squad`, memperluas synthetic formula data menjadi 30 template, mengubah sekitar 20% baris training menjadi gaya slide, serta menandai answer/formula dengan special tokens.
4. Training menyimpan model dan tokenizer ke `/kaggle/working/mt5-base-aqg-final`. Unduh folder hasil itu dari Kaggle sebagai arsip sebelum menjalankan pipeline mT5 di komputer lokal.

Jika `accelerate.notebook_launcher` menyatakan CUDA sudah diinisialisasi di proses utama, restart sesi Kaggle lalu jalankan cell dari awal; jangan jalankan cell inferensi yang memanggil CUDA sebelum cell launcher.

## B. Install pipeline lokal

Di Python environment yang digunakan untuk pipeline:

```bash
python -m pip install -r requirements_local.txt
```

`torch` dipasang dengan build yang sesuai dengan mesin lokal. Untuk demo fallback, parser dependencies tidak dibutuhkan selama memakai input contoh built-in; untuk PDF/PPTX/DOCX, pasang library parser yang sesuai.

## C. Jalankan demo dua bahasa

```bash
python aqg_local_pipeline.py --demo --no-model --output aqg_demo_output.json
```

Mode ini sengaja tidak mengklaim memakai mT5. Tanpa bobot, pertanyaan disusun sebagai cloze fallback dan setiap baris ditandai `question_source="template_fallback"`.

## D. Jalankan dengan bobot mT5 lokal

```bash
python aqg_local_pipeline.py --input materi_kuliah.pdf --model-dir ./mt5-base-aqg-final --output kuis.json
```

Ganti input dengan file `.pptx` atau `.docx` juga bisa. Folder `--model-dir` harus memuat bobot fine-tuning, `config.json`, dan seluruh file tokenizer hasil penyimpanan notebook, termasuk special tokens `<answer>`, `</answer>`, `<formula>`, dan `</formula>`.

Jika folder model ada di Kaggle, itu belum otomatis tersedia di komputer lokal. Unduh artefak model dari Kaggle terlebih dahulu.

## E. Unit test

```bash
python -m unittest -v test_aqg_local_pipeline.py
```

## Batasan yang harus dipahami

- PDF parser mengekstrak teks yang tersedia di PDF. PDF hasil scan tanpa lapisan teks tidak akan terbaca; OCR belum termasuk pipeline ini.
- Formula detection, OMML conversion, language detection fallback, answer extraction, dan distractor bersifat heuristik. Bentuk persamaan yang sangat kompleks atau tata letak dokumen yang buruk tetap memerlukan pemeriksaan manual.
- Tiga distractor yang dihasilkan secara sintaktis belum tentu menjadi pengecoh pedagogis yang baik. Periksa kesalahan konsep sebelum kuis dipakai untuk penilaian formal.
- Metrik ROUGE/BLEU tidak membuktikan pertanyaan benar secara faktual atau menjawab answer yang dituju. Audit sensitivitas jawaban juga merupakan diagnostic, bukan jaminan kualitas.
