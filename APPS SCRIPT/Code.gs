
/**
 * GENERATOR SOAL OTOMATIS
 * Versi 2.0
 * Fitur:
 * 1. Import soal Word (.docx)
 * 2. Membaca DESKRIPSI multiline
 * 3. Membuat Google Form otomatis
 * 4. Mendukung PG, PGK, MJ, IS, UR
 * 5. Menggunakan BANK_SOAL kolom A-T
 */

// ==================================================
// 1. MENU SPREADSHEET
// ==================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("GENERATOR SOAL")
    .addItem("Import Soal dari Word", "importSoalWord")
    .addItem("Buat Google Form", "buatGoogleForm")
    .addItem("Tambah Soal ke Form yang Ada", "tambahSoalKeForm")
    .addSeparator()
    .addItem("Validasi Bank Soal", "validasiBankSoal")
    .addItem("Edit Soal", "editSoal")
    .addItem("Hapus Soal", "hapusSoal")
    .addItem("Deteksi Pertanyaan Duplikat", "deteksiDuplikatSoal")
    .addItem("Cari dan Filter Bank Soal", "cariFilterBankSoal")
    .addItem("Reset Filter Bank Soal", "resetFilterBankSoal")
    .addSeparator()
    .addItem("Analisis Hasil Ujian", "analisisHasilUjian")
    .addItem("Analisis Per Soal", "analisisPerSoal")
    .addItem("Ringkasan Analisis Ujian", "analisisRingkasanUjian")
    .addItem("Buat Dashboard Ujian", "buatDashboardUjian")
    .addToUi();
}



// ==================================================
// 2. MEMBACA BLOK SOAL DARI WORD
// ==================================================

function parseSoalBlock_(block) {

  const data = {};

  const recognizedLabels = new Set([
    "ID",
    "JENIS",
    "MAPEL",
    "KELAS",
    "MATERI",
    "TP",
    "BLOOM",
    "LEVEL",
    "DESKRIPSI",
    "PERTANYAAN",
    "OPSI_A",
    "OPSI_B",
    "OPSI_C",
    "OPSI_D",
    "KUNCI",
    "JAWABAN_DITERIMA",
    "PASANGAN",
    "RUBRIK",
    "POIN",
    "PEMBAHASAN"
  ]);

  const content = block
    .replace("[SOAL]", "")
    .replace("[/SOAL]", "")
    .trim();

  let currentKey = "";

  content.split(/\r?\n/).forEach(line => {

    const match = line.match(/^([A-Z_]+):\s*(.*)$/);

    if (match && recognizedLabels.has(match[1])) {

      currentKey = match[1];
      data[currentKey] = match[2].trim();

    } else if (
      currentKey === "DESKRIPSI" ||
      currentKey === "PERTANYAAN"
    ) {

      data[currentKey] =
        (data[currentKey] ? data[currentKey] + "\n" : "") +
        line;

    }

  });

  if (data.DESKRIPSI) {
    data.DESKRIPSI = data.DESKRIPSI.trim();
  }

  if (data.PERTANYAAN) {
    data.PERTANYAAN = data.PERTANYAAN.trim();
  }

  return data;
}


// ==================================================
// 3. IMPORT SOAL DARI WORD KE BANK_SOAL
// ==================================================

function importSoalWord() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "IMPORT SOAL WORD",
    "Masukkan ID file Word dari Google Drive:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const fileId = response.getResponseText().trim();

  if (!fileId) {
    ui.alert("ID file tidak boleh kosong.");
    return;
  }

  let convertedFileId = null;

  try {

    const wordFile = DriveApp.getFileById(fileId);

    if (!wordFile.getName().toLowerCase().endsWith(".docx")) {
      throw new Error("File harus berformat .docx");
    }

    // Mengonversi Word menjadi Google Docs
    const converted = Drive.Files.create(
      {
        name: "TEMP_IMPORT_" + wordFile.getName(),
        mimeType: "application/vnd.google-apps.document"
      },
      wordFile.getBlob(),
      {
        fields: "id,name"
      }
    );

    convertedFileId = converted.id;

    // Membaca isi dokumen
    const document = DocumentApp.openById(convertedFileId);
    const text = document.getBody().getText();

    const questions = text.match(
      /\[SOAL\][\s\S]*?\[\/SOAL\]/g
    ) || [];

    if (questions.length === 0) {
      throw new Error(
        "Tidak ditemukan blok soal. Periksa format [SOAL]."
      );
    }

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("BANK_SOAL");

    if (!sheet) {
      throw new Error("Sheet BANK_SOAL tidak ditemukan.");
    }

    const existingIds = sheet.getLastRow() > 1
      ? sheet.getRange(
          2, 1, sheet.getLastRow() - 1, 1
        ).getValues().flat().map(x => String(x).trim())
      : [];

    const newRows = [];
    const errors = [];
    const importedIds = new Set(existingIds);

    questions.forEach((block, index) => {

      const data = parseSoalBlock_(block);

      const type = String(data.JENIS || "").trim().toUpperCase();
      const id = String(data.ID || "").trim();

      // Data wajib
      const required = [
        "ID",
        "JENIS",
        "MAPEL",
        "KELAS",
        "MATERI",
        "TP",
        "BLOOM",
        "LEVEL",
        "PERTANYAAN",
        "POIN"
      ];

      const missing = required.filter(
        key => !data[key]
      );

      if (missing.length > 0) {
        errors.push(
          "Soal " + (index + 1) +
          ": Data kosong: " + missing.join(", ")
        );
        return;
      }

      if (!["PG", "PGK", "MJ", "IS", "UR"].includes(type)) {
        errors.push("Soal " + id + ": Jenis soal tidak valid.");
        return;
      }

      if (!/^C[1-6]$/.test(data.BLOOM)) {
        errors.push("Soal " + id + ": Level Bloom tidak valid.");
        return;
      }

      if (!/^L[1-3]$/.test(data.LEVEL)) {
        errors.push("Soal " + id + ": Level kognitif tidak valid.");
        return;
      }

      if (!/^\d+$/.test(data.POIN)) {
        errors.push(
          "Soal " + id + ": Poin harus berupa bilangan bulat."
        );
        return;
      }

      if (importedIds.has(id)) {
        errors.push("Soal " + id + ": ID sudah digunakan.");
        return;
      }

      // Pemeriksaan khusus PG dan PGK
      if (type === "PG" || type === "PGK") {

        if (
          !data.OPSI_A ||
          !data.OPSI_B ||
          !data.OPSI_C ||
          !data.OPSI_D
        ) {
          errors.push("Soal " + id + ": Opsi A-D wajib diisi.");
          return;
        }

        if (!data.KUNCI) {
          errors.push("Soal " + id + ": Kunci wajib diisi.");
          return;
        }

        const keys = data.KUNCI.split(",").map(
          x => x.trim().toUpperCase()
        );

        if (
          keys.some(x => !["A", "B", "C", "D"].includes(x)) ||
          new Set(keys).size !== keys.length ||
          (type === "PG" && keys.length !== 1) ||
          (type === "PGK" && keys.length < 2)
        ) {
          errors.push("Soal " + id + ": Format kunci tidak valid.");
          return;
        }
      }

      // Pemeriksaan menjodohkan
      if (type === "MJ" && !data.PASANGAN) {
        errors.push("Soal " + id + ": Pasangan wajib diisi.");
        return;
      }

      // Pemeriksaan isian singkat
      if (type === "IS" && !data.JAWABAN_DITERIMA) {
        errors.push("Soal " + id + ": Jawaban diterima wajib diisi.");
        return;
      }

      // Pemeriksaan uraian
      if (type === "UR" && !data.RUBRIK) {
        errors.push("Soal " + id + ": Rubrik wajib diisi.");
        return;
      }

      // Memasukkan data ke kolom A-T
      newRows.push([
        data.ID,
        type,
        data.MAPEL,
        data.KELAS,
        data.MATERI,
        data.TP,
        data.BLOOM,
        data.LEVEL,
        data.PERTANYAAN,
        data.OPSI_A || "",
        data.OPSI_B || "",
        data.OPSI_C || "",
        data.OPSI_D || "",
        data.KUNCI || "",
        data.JAWABAN_DITERIMA || "",
        data.PASANGAN || "",
        data.RUBRIK || "",
        Number(data.POIN),
        data.PEMBAHASAN || "",
        data.DESKRIPSI || ""
      ]);

      importedIds.add(id);

    });

    // Menambahkan hasil impor
    if (newRows.length > 0) {

      sheet.getRange(
        sheet.getLastRow() + 1,
        1,
        newRows.length,
        20
      ).setValues(newRows);

    }

    let message =
      "IMPORT SELESAI\n\n" +
      "Soal berhasil: " + newRows.length +
      "\nSoal bermasalah: " + errors.length;

    if (errors.length > 0) {
      message += "\n\nRINCIAN KESALAHAN:\n" +
        errors.join("\n");
    }

    ui.alert(message);

  } catch (error) {

    ui.alert(
      "IMPORT GAGAL\n\n" + error.message
    );

  } finally {

    // Menghapus dokumen sementara
    if (convertedFileId) {
      try {
        DriveApp.getFileById(convertedFileId).setTrashed(true);
      } catch (cleanupError) {
        console.error(cleanupError);
      }
    }

  }

}


// ==================================================
// 4. MEMBUAT GOOGLE FORM
// ==================================================


function buatGoogleFormDariHTML(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANK_SOAL");

  if (!data) {
    throw new Error("Data dari HTML tidak diterima.");
  }

  if (!sheet) {
    throw new Error("Sheet BANK_SOAL tidak ditemukan.");
  }

  if (sheet.getLastRow() < 2) {
    throw new Error("BANK_SOAL belum memiliki data soal.");
  }

  // Mengambil pengaturan dari HTML
  const formTitle = String(data.formTitle || "").trim();
  const formDescription = String(data.formDescription || "").trim();
  const shuffleQuestions =
    data.shuffleQuestions === true ||
    String(data.shuffleQuestions).toLowerCase() === "true";

  if (!formTitle) {
    throw new Error("Judul ujian tidak boleh kosong.");
  }

  // Membaca kolom A-U
  const values = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 21
  ).getValues();

  const allQuestions = values.filter(row =>
    String(row[0]).trim() !== "" &&
    String(row[8]).trim() !== ""
  );

  if (allQuestions.length === 0) {
    throw new Error("Tidak ada soal valid di BANK_SOAL.");
  }

  let questions = [];

  // METODE 1: BERDASARKAN ID
  if (data.method === "id") {

    const selectionText = String(data.selectionText || "").trim();

    if (!selectionText) {
      throw new Error("ID soal belum diisi.");
    }

    const selectedIds = selectionText === "*"
      ? null
      : [...new Set(
          selectionText.split(",")
            .map(id => id.trim())
            .filter(Boolean)
        )];

    if (selectedIds !== null) {

      if (selectedIds.length === 0) {
        throw new Error("Format ID soal tidak valid.");
      }

      const existingIds = new Set(
        allQuestions.map(row => String(row[0]).trim())
      );

      const notFound = selectedIds.filter(
        id => !existingIds.has(id)
      );

      if (notFound.length > 0) {
        throw new Error(
          "ID SOAL TIDAK DITEMUKAN:\n" +
          notFound.join("\n")
        );
      }
    }

    questions = selectedIds === null
      ? [...allQuestions]
      : allQuestions.filter(row =>
          selectedIds.includes(String(row[0]).trim())
        );

  // METODE 2: ACAK
  } else if (data.method === "random") {

    const numberOfQuestions = Number(
      data.numberOfQuestions || data.randomCount
    );

    if (
      !Number.isInteger(numberOfQuestions) ||
      numberOfQuestions < 1 ||
      numberOfQuestions > allQuestions.length
    ) {
      throw new Error(
        "Jumlah soal tidak valid. Soal tersedia: " +
        allQuestions.length
      );
    }

    questions = [...allQuestions];

    for (let i = questions.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [questions[i], questions[j]] =
        [questions[j], questions[i]];
    }

    questions = questions.slice(0, numberOfQuestions);

  // METODE 3: BERDASARKAN KRITERIA
  } else if (data.method === "criteria") {

    const subject = String(data.subject || "").trim();
    const grade = String(data.grade || "").trim();
    const material = String(data.material || "").trim();
    const bloom = String(data.bloom || "").trim();
    const cognitive = String(data.cognitive || "").trim();

    const filteredQuestions = allQuestions.filter(row =>
      (!subject ||
        String(row[2] || "").trim().toLowerCase() === subject.toLowerCase()) &&

      (!grade ||
        String(row[3] || "").trim().toLowerCase() === grade.toLowerCase()) &&

      (!material ||
        String(row[4] || "").trim().toLowerCase().includes(material.toLowerCase())) &&

      (!bloom ||
        String(row[6] || "").trim().toLowerCase() === bloom.toLowerCase()) &&

      (!cognitive ||
        String(row[7] || "").trim().toLowerCase() === cognitive.toLowerCase())
    );

    if (filteredQuestions.length === 0) {
      throw new Error(
        "Tidak ada soal yang sesuai dengan kriteria tersebut."
      );
    }

    const numberOfQuestions = Number(
      data.numberOfQuestions || data.criteriaCount
    );

    if (
      !Number.isInteger(numberOfQuestions) ||
      numberOfQuestions < 1 ||
      numberOfQuestions > filteredQuestions.length
    ) {
      throw new Error(
        "Jumlah soal tidak valid. Soal sesuai kriteria: " +
        filteredQuestions.length
      );
    }

    questions = [...filteredQuestions];

    for (let i = questions.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [questions[i], questions[j]] =
        [questions[j], questions[i]];
    }

    questions = questions.slice(0, numberOfQuestions);

  } else {
    throw new Error("Metode pemilihan soal tidak dikenal.");
  }

  if (questions.length === 0) {
    throw new Error("Tidak ditemukan soal yang dipilih.");
  }

  const plan = [];
  const errors = [];

  // MEMVALIDASI SEMUA SOAL SEBELUM MEMBUAT FORM
  questions.forEach(row => {

    const id = String(row[0]).trim();
    const type = String(row[1]).trim().toUpperCase();
    const question = String(row[8]).trim();
    const points = Number(row[17]);

    const description = String(row[19] || "").trim();
    const imageId = String(row[20] || "").trim();

    const material = String(row[4] || "").trim();
    const objective = String(row[5] || "").trim();
    const bloom = String(row[6] || "").trim();
    const cognitive = String(row[7] || "").trim();

    if (!Number.isInteger(points) || points < 0) {
      errors.push(id + ": POIN harus berupa bilangan bulat minimal 0.");
      return;
    }

    const options = [
      String(row[9] || "").trim(),
      String(row[10] || "").trim(),
      String(row[11] || "").trim(),
      String(row[12] || "").trim()
    ];

    const key = String(row[13] || "")
      .trim()
      .split(",")
      .map(x => x.trim().toUpperCase())
      .filter(Boolean);

    // SOAL PILIHAN GANDA
    if (type === "PG") {

      if (options.some(x => !x)) {
        errors.push(id + ": Opsi A-D belum lengkap.");
        return;
      }

      if (
        key.length !== 1 ||
        !["A", "B", "C", "D"].includes(key[0])
      ) {
        errors.push(id + ": Kunci PG tidak valid.");
        return;
      }

      plan.push({
        id, type, question, points,
        material, objective, bloom, cognitive,
        description, options, key, imageId
      });

    // SOAL PILIHAN GANDA KOMPLEKS
    } else if (type === "PGK") {

      if (options.some(x => !x)) {
        errors.push(id + ": Opsi A-D belum lengkap.");
        return;
      }

      if (
        key.length < 2 ||
        key.some(x => !["A", "B", "C", "D"].includes(x)) ||
        new Set(key).size !== key.length
      ) {
        errors.push(id + ": Kunci PGK tidak valid.");
        return;
      }

      plan.push({
        id, type, question, points,
        material, objective, bloom, cognitive,
        description, options, key, imageId
      });

    // SOAL MENJODOHKAN
    } else if (type === "MJ") {

      const rawPairs = String(row[15] || "").trim();

      if (!rawPairs) {
        errors.push(id + ": Data PASANGAN kosong.");
        return;
      }

      const pairs = rawPairs.split(";").map(pair => {
        const parts = pair.split("=");

        if (parts.length < 2) {
          return null;
        }

        return {
          left: parts[0].trim(),
          right: parts.slice(1).join("=").trim()
        };
      });

      if (pairs.some(p => !p || !p.left || !p.right)) {
        errors.push(
          id + ": Format pasangan harus kiri=kanan; kiri=kanan."
        );
        return;
      }

      if (points < pairs.length) {
        errors.push(
          id + ": POIN minimal harus sama dengan jumlah pasangan (" +
          pairs.length + ")."
        );
        return;
      }

      const rightOptions = [...new Set(
        pairs.map(p => p.right)
      )];

      if (rightOptions.length !== pairs.length) {
        errors.push(
          id + ": Jawaban kanan pada pasangan tidak boleh sama."
        );
        return;
      }

      plan.push({
        id, type, question, points,
        material, objective, bloom, cognitive,
        description, pairs, rightOptions
      });

    // SOAL ISIAN SINGKAT
    } else if (type === "IS") {

      const accepted = String(row[14] || "").trim();

      if (!accepted) {
        errors.push(id + ": JAWABAN_DITERIMA kosong.");
        return;
      }

      plan.push({
        id, type, question, points,
        material, objective, bloom, cognitive,
        description, accepted
      });

    // SOAL URAIAN
    } else if (type === "UR") {

      const rubric = String(row[16] || "").trim();

      if (!rubric) {
        errors.push(id + ": RUBRIK kosong.");
        return;
      }

      plan.push({
        id, type, question, points,
        material, objective, bloom, cognitive,
        description, rubric
      });

    } else {

      errors.push(id + ": Jenis soal tidak dikenal.");

    }

  });

  if (plan.length === 0) {
    throw new Error(
      "Tidak ada soal yang valid untuk dibuat menjadi Google Form."
    );
  }

  if (errors.length > 0) {
    throw new Error(
      "GOOGLE FORM BELUM DIBUAT:\n\n" +
      errors.join("\n")
    );
  }

  let form = null;

  try {

    // MEMBUAT GOOGLE FORM
    form = FormApp.create(formTitle);

    form.setIsQuiz(true);
    form.setDescription(formDescription);
    form.setConfirmationMessage(
      "Jawaban kamu sudah berhasil dikirim."
    );
    form.setProgressBar(true);
    form.setShuffleQuestions(shuffleQuestions);

    // IDENTITAS PESERTA
    form.addTextItem()
      .setTitle("Nama Lengkap")
      .setRequired(true);

    form.addTextItem()
      .setTitle("Kelas")
      .setRequired(true);

    form.addTextItem()
      .setTitle("Nomor Absen")
      .setRequired(true);

    // MEMISAHKAN HALAMAN IDENTITAS DAN HALAMAN UJIAN
    form.addPageBreakItem()
      .setTitle("SOAL UJIAN")
      .setHelpText("Kerjakan semua pertanyaan dengan teliti.");


    // Menyiapkan penampung pemetaan soal
    const mappingRows = [];
    let nomorItem = 0;

    // Membuat soal
    plan.forEach(q => {

      // Menambahkan gambar jika tersedia
      if (q.imageId) {
        const imageFile = DriveApp.getFileById(q.imageId);

        form.addImageItem()
          .setTitle("Gambar soal " + q.id)
          .setImage(imageFile.getBlob())
          .setAlignment(FormApp.Alignment.CENTER)
          .setWidth(500);
      }

      // SOAL PILIHAN GANDA
      if (q.type === "PG") {

        const item = form.addMultipleChoiceItem();
        item.setTitle(q.question);

        if (q.description) {
          item.setHelpText(q.description);
        }

        const choices = q.options.map((option, index) => {
          const letter = String.fromCharCode(65 + index);
          return item.createChoice(option, q.key[0] === letter);
        });

        item.setChoices(choices);
        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          JSON.stringify(q.key), "", "",
          q.material, q.bloom, q.cognitive,
          "DINILAI_OTOMATIS"
        ]);

      // SOAL PILIHAN GANDA KOMPLEKS
      } else if (q.type === "PGK") {

        const item = form.addCheckboxItem();
        item.setTitle(q.question);

        if (q.description) {
          item.setHelpText(q.description);
        }

        const choices = q.options.map((option, index) => {
          const letter = String.fromCharCode(65 + index);
          return item.createChoice(option, q.key.includes(letter));
        });

        item.setChoices(choices);
        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          JSON.stringify(
            q.key.map(letter =>
              q.options[letter.charCodeAt(0) - 65]
            )
          ),
          "", "",
          q.material, q.bloom, q.cognitive,
          "DINILAI_OTOMATIS"
        ]);

      // SOAL MENJODOHKAN
      } else if (q.type === "MJ") {

        const quotient = Math.floor(q.points / q.pairs.length);
        const remainder = q.points % q.pairs.length;

        q.pairs.forEach((pair, index) => {

          const item = form.addMultipleChoiceItem();

          item.setTitle(
            q.id + " - " + q.question +
            "\nPasangan " + (index + 1) +
            ": " + pair.left
          );

          if (q.description) {
            item.setHelpText(q.description);
          }

          const choices = q.rightOptions.map(option =>
            item.createChoice(option, option === pair.right)
          );

          item.setChoices(choices);

          const pairPoints =
            quotient + (index < remainder ? 1 : 0);

          item.setPoints(pairPoints);
          item.setRequired(true);

          nomorItem++;

          mappingRows.push([
            q.id, item.getId(), nomorItem, index + 1,
            q.type, q.question + " - " + pair.left,
            pairPoints, JSON.stringify(pair.right), "", "",
            q.material, q.bloom, q.cognitive,
            "DINILAI_OTOMATIS"
          ]);

        });

      // SOAL ISIAN SINGKAT
      } else if (q.type === "IS") {

        const item = form.addTextItem();
        item.setTitle(q.question);

        item.setHelpText(
          (q.description ? q.description + "\n\n" : "") +
          "Jawaban akan diperiksa oleh guru."
        );

        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          "", JSON.stringify([q.accepted]), "",
          q.material, q.bloom, q.cognitive,
          "PERLU_PENILAIAN_GURU"
        ]);

      // SOAL URAIAN
      } else if (q.type === "UR") {

        const item = form.addParagraphTextItem();
        item.setTitle(q.question);

        item.setHelpText(
          (q.description ? q.description + "\n\n" : "") +
          "Jawaban akan diperiksa berdasarkan rubrik guru."
        );

        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          "", "", JSON.stringify(q.rubric),
          q.material, q.bloom, q.cognitive,
          "PERLU_PENILAIAN_GURU"
        ]);

      }

    }); // Menutup plan.forEach()

    // Membuat kisi-kisi berdasarkan soal yang digunakan
    buatKisiKisi_(plan);

    // Memastikan sheet pemetaan tersedia
    const pemetaanSheet = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");
    if (!pemetaanSheet) {
      throw new Error("Sheet PEMETAAN_SOAL_UJIAN tidak ditemukan.");
    }

    // Memastikan sheet daftar ujian tersedia
    const daftarUjian = ss.getSheetByName("DAFTAR_UJIAN");
    if (!daftarUjian) {
      throw new Error("Sheet DAFTAR_UJIAN tidak ditemukan.");
    }

    // Mencatat daftar sheet sebelum menghubungkan respons
    const sheetNamesBefore = ss.getSheets().map(
      s => s.getName()
    );

    // Menghubungkan respons ke spreadsheet
    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      ss.getId()
    );

    SpreadsheetApp.flush();

    // Mendeteksi sheet respons yang baru dibuat
    const newResponseSheet = ss.getSheets().find(
      s => !sheetNamesBefore.includes(s.getName())
    );

    const responseSheetName = newResponseSheet
      ? newResponseSheet.getName()
      : "";

    // Membuat ID ujian unik
    const idUjian = "UJIAN-" + Utilities.getUuid();

    // Menyimpan pemetaan setiap soal ke spreadsheet
    const mappingData = mappingRows.map(row => [
      idUjian,
      ...row
    ]);

    if (mappingData.length > 0) {
      pemetaanSheet.getRange(
        pemetaanSheet.getLastRow() + 1,
        1,
        mappingData.length,
        15
      ).setValues(mappingData);
    }

    // Menyimpan identitas ujian
    daftarUjian.appendRow([
      idUjian,
      formTitle,
      form.getId(),
      form.getPublishedUrl(),
      ss.getId(),
      responseSheetName,
      new Date(),
      "",
      plan.length,
      "DIBUAT"
    ]);

    // Mengirim hasil pembuatan Form ke HTML
    return {
      success: true,
      count: plan.length,
      editUrl: form.getEditUrl(),
      publishedUrl: form.getPublishedUrl()
    };

  } catch (error) {

    throw new Error(
      "GAGAL MEMBUAT GOOGLE FORM:\n\n" +
      error.message +
      (form
        ? "\n\nForm mungkin sudah dibuat sebagian. Periksa Google Drive."
        : "")
    );

  }
}

function buatKisiKisi_(questions) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("KISI_KISI");

  if (!sheet) {
    throw new Error("Sheet KISI_KISI tidak ditemukan.");
  }

  // Membersihkan kisi-kisi sebelumnya
  sheet.clear();

  // Judul laporan
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1")
    .setValue("KISI-KISI UJIAN")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Judul tabel
  const headers = [
    "NO.",
    "ID SOAL",
    "MATERI",
    "TUJUAN PEMBELAJARAN",
    "LEVEL BLOOM",
    "LEVEL KOGNITIF",
    "JENIS SOAL",
    "POIN"
  ];

  
  sheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#d9ead3");

  // Menyiapkan rincian soal
  const rows = questions.map((q, index) => [
    index + 1,
    q.id,
    q.material || "",
    q.objective || "",
    q.bloom || "",
    q.cognitive || "",
    q.type,
    q.points
  ]);

  // Menuliskan rincian soal
  if (rows.length > 0) {
    sheet.getRange(4, 1, rows.length, 8)
      .setValues(rows);
  }

}

function tambahSoalKeForm() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let form = null;
  let itemsAdded = 0;

  try {

    // MEMINTA LINK GOOGLE FORM
    const response = ui.prompt(
      "TAMBAH SOAL KE FORM",
      "Masukkan link Google Form yang ingin ditambahkan soal:",
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    const formUrl = response.getResponseText().trim();
    const match = formUrl.match(/[-\w]{25,}/);

    if (!match) {
      throw new Error("Link atau ID Google Form tidak valid.");
    }

    // MEMBUKA GOOGLE FORM
    form = FormApp.openById(match[0]);

    if (!form.isQuiz()) {
      throw new Error(
        "Form tujuan belum diaktifkan sebagai kuis."
      );
    }

    // MENCARI ID_UJIAN
    const examSheet = ss.getSheetByName("DAFTAR_UJIAN");

    if (!examSheet) {
      throw new Error("Sheet DAFTAR_UJIAN tidak ditemukan.");
    }

    const examData = examSheet.getDataRange().getValues();

    const examRow = examData.slice(1).find(row =>
      String(row[2]).trim() === match[0]
    );

    if (!examRow) {
      throw new Error(
        "Google Form belum terdaftar di DAFTAR_UJIAN."
      );
    }

    const idUjian = String(examRow[0]).trim();

    // MEMERIKSA SHEET PEMETAAN
    const mappingSheet = ss.getSheetByName(
      "PEMETAAN_SOAL_UJIAN"
    );

    if (!mappingSheet) {
      throw new Error(
        "Sheet PEMETAAN_SOAL_UJIAN tidak ditemukan."
      );
    }

    // MENYIAPKAN LOG SOAL
    let logSheet = ss.getSheetByName("LOG_SOAL_FORM");

    if (!logSheet) {
      logSheet = ss.insertSheet("LOG_SOAL_FORM");

      logSheet.appendRow([
        "FORM_ID",
        "ID_SOAL",
        "TANGGAL_DITAMBAHKAN"
      ]);
    }

    const logData = logSheet.getDataRange().getValues();

    const loggedIds = new Set(
      logData.slice(1)
        .filter(row => String(row[0]).trim() === match[0])
        .map(row => String(row[1]).trim())
    );

    // MEMBACA BANK SOAL
    const sheet = ss.getSheetByName("BANK_SOAL");

    if (!sheet) {
      throw new Error("Sheet BANK_SOAL tidak ditemukan.");
    }

    if (sheet.getLastRow() < 2) {
      throw new Error("BANK_SOAL belum memiliki data soal.");
    }

    const values = sheet.getRange(
      2, 1, sheet.getLastRow() - 1, 21
    ).getValues();

    const allQuestions = values.filter(row =>
      String(row[0]).trim() !== "" &&
      String(row[8]).trim() !== ""
    );

    if (allQuestions.length === 0) {
      throw new Error("Tidak ada soal dalam BANK_SOAL.");
    }

    // MEMILIH METODE PENGAMBILAN SOAL
    const methodChoice = ui.alert(
      "METODE PEMILIHAN SOAL",
      "Pilih metode pengambilan soal:\n\n" +
      "YA = Berdasarkan ID soal\n" +
      "TIDAK = Acak dari seluruh BANK_SOAL\n" +
      "BATAL = Berdasarkan kriteria tertentu",
      ui.ButtonSet.YES_NO_CANCEL
    );

    let questions = [];

    // METODE 1: BERDASARKAN ID
    if (methodChoice === ui.Button.YES) {

      const response = ui.prompt(
        "PILIH SOAL",
        "Masukkan ID soal dipisahkan koma.\n" +
        "Contoh: SOAL001, SOAL002\n\n" +
        "Ketik * untuk semua soal.",
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() !== ui.Button.OK) return;

      const selectionText = response.getResponseText().trim();

      if (!selectionText) {
        throw new Error("ID soal belum diisi.");
      }

      const selectedIds = selectionText === "*"
        ? null
        : [...new Set(
            selectionText.split(",")
              .map(id => id.trim())
              .filter(Boolean)
          )];

      if (selectedIds !== null) {

        const existingIds = new Set(
          allQuestions.map(row => String(row[0]).trim())
        );

        const notFound = selectedIds.filter(
          id => !existingIds.has(id)
        );

        if (notFound.length > 0) {
          throw new Error(
            "ID SOAL TIDAK DITEMUKAN:\n\n" +
            notFound.join("\n")
          );
        }
      }

      questions = selectedIds === null
        ? [...allQuestions]
        : allQuestions.filter(row =>
            selectedIds.includes(String(row[0]).trim())
          );

    // METODE 2: ACAK
    } else if (methodChoice === ui.Button.NO) {

      const response = ui.prompt(
        "JUMLAH SOAL ACAK",
        "Jumlah soal tersedia: " + allQuestions.length +
        "\nBerapa soal yang ingin ditambahkan?",
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() !== ui.Button.OK) return;

      const numberOfQuestions = Number(
        response.getResponseText().trim()
      );

      if (
        !Number.isInteger(numberOfQuestions) ||
        numberOfQuestions < 1 ||
        numberOfQuestions > allQuestions.length
      ) {
        throw new Error("Jumlah soal tidak valid.");
      }

      questions = [...allQuestions];

      for (let i = questions.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [questions[i], questions[j]] =
          [questions[j], questions[i]];
      }

      questions = questions.slice(0, numberOfQuestions);

    // METODE 3: KRITERIA
    } else if (methodChoice === ui.Button.CANCEL) {

      const askCriterion = (title, message) => {
        const response = ui.prompt(
          title,
          message,
          ui.ButtonSet.OK_CANCEL
        );

        if (response.getSelectedButton() !== ui.Button.OK) {
          return null;
        }

        return response.getResponseText().trim();
      };

      const subject = askCriterion(
        "MATA PELAJARAN",
        "Masukkan mata pelajaran. Kosongkan jika tidak digunakan."
      );
      if (subject === null) return;

      const grade = askCriterion(
        "KELAS",
        "Masukkan kelas. Kosongkan jika tidak digunakan."
      );
      if (grade === null) return;

      const material = askCriterion(
        "MATERI",
        "Masukkan kata kunci materi. Kosongkan jika tidak digunakan."
      );
      if (material === null) return;

      const bloom = askCriterion(
        "LEVEL BLOOM",
        "Masukkan C1 sampai C6. Kosongkan jika tidak digunakan."
      );
      if (bloom === null) return;

      const cognitive = askCriterion(
        "LEVEL KOGNITIF",
        "Masukkan L1, L2, atau L3. Kosongkan jika tidak digunakan."
      );
      if (cognitive === null) return;

      const filteredQuestions = allQuestions.filter(row =>
        (!subject ||
          String(row[2] || "").trim().toLowerCase() === subject.toLowerCase()) &&

        (!grade ||
          String(row[3] || "").trim().toLowerCase() === grade.toLowerCase()) &&

        (!material ||
          String(row[4] || "").trim().toLowerCase().includes(material.toLowerCase())) &&

        (!bloom ||
          String(row[6] || "").trim().toLowerCase() === bloom.toLowerCase()) &&

        (!cognitive ||
          String(row[7] || "").trim().toLowerCase() === cognitive.toLowerCase())
      );

      if (filteredQuestions.length === 0) {
        throw new Error(
          "Tidak ada soal yang sesuai dengan kriteria."
        );
      }

      const response = ui.prompt(
        "JUMLAH SOAL",
        "Soal yang sesuai: " + filteredQuestions.length +
        "\nBerapa soal yang ingin ditambahkan secara acak?",
        ui.ButtonSet.OK_CANCEL
      );

      if (response.getSelectedButton() !== ui.Button.OK) return;

      const numberOfQuestions = Number(
        response.getResponseText().trim()
      );

      if (
        !Number.isInteger(numberOfQuestions) ||
        numberOfQuestions < 1 ||
        numberOfQuestions > filteredQuestions.length
      ) {
        throw new Error("Jumlah soal tidak valid.");
      }

      questions = [...filteredQuestions];

      for (let i = questions.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [questions[i], questions[j]] =
          [questions[j], questions[i]];
      }

      questions = questions.slice(0, numberOfQuestions);

    } else {
      return;
    }

    if (questions.length === 0) {
      throw new Error("Tidak ada soal yang dipilih.");
    }

    // MEMERIKSA SOAL DUPLIKAT
    const alreadyAdded = questions
      .map(row => String(row[0]).trim())
      .filter(id => loggedIds.has(id));

    if (alreadyAdded.length > 0) {
      throw new Error(
        "SOAL SUDAH PERNAH DITAMBAHKAN:\n\n" +
        alreadyAdded.join("\n")
      );
    }

    // MEMVALIDASI SOAL
    const plan = [];
    const errors = [];

    questions.forEach(row => {

      const id = String(row[0]).trim();
      const type = String(row[1]).trim().toUpperCase();
      const question = String(row[8]).trim();
      const points = Number(row[17]);

      const description = String(row[19] || "").trim();
      const imageId = String(row[20] || "").trim();

      const material = String(row[4] || "").trim();
      const objective = String(row[5] || "").trim();
      const bloom = String(row[6] || "").trim();
      const cognitive = String(row[7] || "").trim();

      if (!Number.isInteger(points) || points < 0) {
        errors.push(id + ": POIN tidak valid.");
        return;
      }

      const options = [
        String(row[9] || "").trim(),
        String(row[10] || "").trim(),
        String(row[11] || "").trim(),
        String(row[12] || "").trim()
      ];

      const key = String(row[13] || "")
        .trim()
        .split(",")
        .map(x => x.trim().toUpperCase())
        .filter(Boolean);

      // PILIHAN GANDA DAN PILIHAN GANDA KOMPLEKS
      if (type === "PG" || type === "PGK") {

        if (options.some(x => !x)) {
          errors.push(id + ": Opsi A-D belum lengkap.");
          return;
        }

        if (
          (type === "PG" && (
            key.length !== 1 ||
            !["A", "B", "C", "D"].includes(key[0])
          )) ||
          (type === "PGK" && (
            key.length < 2 ||
            key.some(x => !["A", "B", "C", "D"].includes(x)) ||
            new Set(key).size !== key.length
          ))
        ) {
          errors.push(id + ": Kunci jawaban tidak valid.");
          return;
        }

        plan.push({
          id, type, question, points,
          material, objective, bloom, cognitive,
          description, options, key, imageId
        });

      // MENJODOHKAN
      } else if (type === "MJ") {

        const rawPairs = String(row[15] || "").trim();

        if (!rawPairs) {
          errors.push(id + ": Data pasangan kosong.");
          return;
        }

        const pairs = rawPairs.split(";").map(pair => {
          const parts = pair.split("=");

          if (parts.length < 2) return null;

          return {
            left: parts[0].trim(),
            right: parts.slice(1).join("=").trim()
          };
        });

        if (pairs.some(p => !p || !p.left || !p.right)) {
          errors.push(id + ": Format pasangan tidak valid.");
          return;
        }

        const rightOptions = [...new Set(
          pairs.map(p => p.right)
        )];

        if (
          rightOptions.length !== pairs.length ||
          points < pairs.length
        ) {
          errors.push(id + ": Pasangan atau poin tidak valid.");
          return;
        }

        plan.push({
          id, type, question, points,
          material, objective, bloom, cognitive,
          description, pairs, rightOptions, imageId
        });

      // ISIAN SINGKAT
      } else if (type === "IS") {

        const accepted = String(row[14] || "").trim();

        if (!accepted) {
          errors.push(id + ": Jawaban diterima kosong.");
          return;
        }

        plan.push({
          id, type, question, points,
          material, objective, bloom, cognitive,
          description, accepted, imageId
        });

      // URAIAN
      } else if (type === "UR") {

        const rubric = String(row[16] || "").trim();

        if (!rubric) {
          errors.push(id + ": Rubrik kosong.");
          return;
        }

        plan.push({
          id, type, question, points,
          material, objective, bloom, cognitive,
          description, rubric, imageId
        });

      } else {

        errors.push(id + ": Jenis soal tidak dikenal.");

      }

    });

    if (plan.length === 0 || errors.length > 0) {
      throw new Error(
        "VALIDASI SOAL GAGAL:\n\n" +
        (errors.length > 0
          ? errors.join("\n")
          : "Tidak ada soal valid.")
      );
    }

    // KONFIRMASI PENAMBAHAN
    const confirm = ui.alert(
      "KONFIRMASI PENAMBAHAN",
      "Form tujuan: " + form.getTitle() +
      "\nJumlah soal: " + plan.length +
      "\n\nSoal akan ditambahkan di bagian akhir Form." +
      "\nSoal dan respons lama tetap dipertahankan." +
      "\n\nLanjutkan?",
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    // MENENTUKAN NOMOR ITEM BERIKUTNYA
    const existingMapping = mappingSheet.getDataRange().getValues();

    let nomorItem = 0;

    existingMapping.slice(1).forEach(row => {
      if (String(row[0]).trim() === idUjian) {
        nomorItem = Math.max(
          nomorItem,
          Number(row[3]) || 0
        );
      }
    });

    const mappingRows = [];

    // MENAMBAHKAN SOAL KE FORM
    plan.forEach(q => {

      // MENAMBAHKAN GAMBAR
      if (q.imageId) {
        const imageFile = DriveApp.getFileById(q.imageId);

        form.addImageItem()
          .setTitle("Gambar soal " + q.id)
          .setImage(imageFile.getBlob())
          .setAlignment(FormApp.Alignment.CENTER)
          .setWidth(500);
      }

      // PILIHAN GANDA
      if (q.type === "PG") {

        const item = form.addMultipleChoiceItem();
        item.setTitle(q.question);

        if (q.description) {
          item.setHelpText(q.description);
        }

        const choices = q.options.map((option, index) => {
          const letter = String.fromCharCode(65 + index);
          return item.createChoice(option, q.key[0] === letter);
        });

        item.setChoices(choices);
        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          JSON.stringify(q.key), "", "",
          q.material, q.bloom, q.cognitive,
          "DINILAI_OTOMATIS"
        ]);

      // PILIHAN GANDA KOMPLEKS
      } else if (q.type === "PGK") {

        const item = form.addCheckboxItem();
        item.setTitle(q.question);

        if (q.description) {
          item.setHelpText(q.description);
        }

        const choices = q.options.map((option, index) => {
          const letter = String.fromCharCode(65 + index);
          return item.createChoice(option, q.key.includes(letter));
        });

        item.setChoices(choices);
        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          JSON.stringify(
            q.key.map(letter =>
              q.options[letter.charCodeAt(0) - 65]
            )
          ),
          "", "",
          q.material, q.bloom, q.cognitive,
          "DINILAI_OTOMATIS"
        ]);

      // MENJODOHKAN
      } else if (q.type === "MJ") {

        const quotient = Math.floor(
          q.points / q.pairs.length
        );

        const remainder = q.points % q.pairs.length;

        q.pairs.forEach((pair, index) => {

          const item = form.addMultipleChoiceItem();

          item.setTitle(
            q.question +
            "\nPasangan " + (index + 1) +
            ": " + pair.left
          );

          if (q.description) {
            item.setHelpText(q.description);
          }

          const choices = q.rightOptions.map(option =>
            item.createChoice(option, option === pair.right)
          );

          item.setChoices(choices);

          const pairPoints =
            quotient + (index < remainder ? 1 : 0);

          item.setPoints(pairPoints);
          item.setRequired(true);

          nomorItem++;

          mappingRows.push([
            q.id, item.getId(), nomorItem, index + 1,
            q.type, item.getTitle(), pairPoints,
            JSON.stringify(pair.right), "", "",
            q.material, q.bloom, q.cognitive,
            "DINILAI_OTOMATIS"
          ]);

        });

      // ISIAN SINGKAT
      } else if (q.type === "IS") {

        const item = form.addTextItem();
        item.setTitle(q.question);

        item.setHelpText(
          (q.description ? q.description + "\n\n" : "") +
          "Jawaban akan diperiksa oleh guru."
        );

        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          "", JSON.stringify([q.accepted]), "",
          q.material, q.bloom, q.cognitive,
          "PERLU_PENILAIAN_GURU"
        ]);

      // URAIAN
      } else if (q.type === "UR") {

        const item = form.addParagraphTextItem();
        item.setTitle(q.question);

        item.setHelpText(
          (q.description ? q.description + "\n\n" : "") +
          "Jawaban akan diperiksa berdasarkan rubrik guru."
        );

        item.setPoints(q.points);
        item.setRequired(true);

        nomorItem++;

        mappingRows.push([
          q.id, item.getId(), nomorItem, "",
          q.type, q.question, q.points,
          "", "", JSON.stringify(q.rubric),
          q.material, q.bloom, q.cognitive,
          "PERLU_PENILAIAN_GURU"
        ]);

      }

      itemsAdded++;

    });

    // MENYIMPAN PEMETAAN SOAL
    if (mappingRows.length > 0) {

      const finalMappingRows = mappingRows.map(row => [
        idUjian,
        ...row
      ]);

      mappingSheet.getRange(
        mappingSheet.getLastRow() + 1,
        1,
        finalMappingRows.length,
        15
      ).setValues(finalMappingRows);

    }

    // MENYIMPAN LOG SOAL
    const logRows = plan.map(q => [
      match[0],
      q.id,
      new Date()
    ]);

    if (logRows.length > 0) {
      logSheet.getRange(
        logSheet.getLastRow() + 1,
        1,
        logRows.length,
        3
      ).setValues(logRows);
    }

    ui.alert(
      "PENAMBAHAN SOAL BERHASIL\n\n" +
      "Form: " + form.getTitle() +
      "\nJumlah soal ditambahkan: " + plan.length
    );

  } catch (error) {

    ui.alert(
      "PROSES PENAMBAHAN SOAL GAGAL\n\n" +
      error.message +
      (itemsAdded > 0
        ? "\n\nPERHATIAN: Sebagian soal mungkin sudah terlanjur ditambahkan ke Form. Periksa Form dan spreadsheet sebelum mencoba kembali."
        : "")
    );

  }

}


// ==================================================
// 7. VALIDASI BANK SOAL
// ==================================================


function editSoal() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANK_SOAL");

  if (!sheet) {
    ui.alert("Sheet BANK_SOAL tidak ditemukan.");
    return;
  }

  if (sheet.getLastRow() < 2) {
    ui.alert("BANK_SOAL belum memiliki data soal.");
    return;
  }

  // MEMINTA ID SOAL
  const idPrompt = ui.prompt(
    "EDIT SOAL",
    "Masukkan ID_SOAL yang akan diedit:",
    ui.ButtonSet.OK_CANCEL
  );

  if (idPrompt.getSelectedButton() !== ui.Button.OK) return;

  const id = idPrompt.getResponseText().trim();

  if (!id) {
    ui.alert("ID_SOAL tidak boleh kosong.");
    return;
  }

  // MENCARI BARIS SOAL
  const lastRow = sheet.getLastRow();

  const data = sheet.getRange(
    2, 1, lastRow - 1, 21
  ).getValues();

  const matchingIndexes = [];

  data.forEach((row, index) => {
    if (String(row[0]).trim() === id) {
      matchingIndexes.push(index);
    }
  });

  if (matchingIndexes.length === 0) {
    ui.alert("ID_SOAL tidak ditemukan.");
    return;
  }

  if (matchingIndexes.length > 1) {
    ui.alert(
      "ID_SOAL DUPLIKAT\n\n" +
      "Ditemukan lebih dari satu soal dengan ID " + id +
      ". Perbaiki ID terlebih dahulu melalui BANK_SOAL."
    );
    return;
  }

  const rowNumber = matchingIndexes[0] + 2;

  // MEMBACA HEADER
  const headers = sheet.getRange(
    1, 1, 1, 21
  ).getValues()[0];

  // MEMILIH KOLOM
  const fieldPrompt = ui.prompt(
    "PILIH KOLOM",
    "Masukkan nomor kolom yang ingin diedit:\n\n" +
    headers.map((header, index) =>
      (index + 1) + ". " + header
    ).join("\n") +
    "\n\nID_SOAL (kolom 1) tidak dapat diubah.",
    ui.ButtonSet.OK_CANCEL
  );

  if (fieldPrompt.getSelectedButton() !== ui.Button.OK) return;

  const column = Number(
    fieldPrompt.getResponseText().trim()
  );

  if (
    !Number.isInteger(column) ||
    column < 2 ||
    column > 21
  ) {
    ui.alert(
      "Nomor kolom tidak valid. Pilih kolom 2 sampai 21."
    );
    return;
  }

  // MEMBACA NILAI LAMA
  const currentRow = sheet.getRange(
    rowNumber, 1, 1, 21
  ).getValues()[0];

  const oldValue = sheet.getRange(
    rowNumber, column
  ).getDisplayValue();

  // MEMINTA NILAI BARU
  const valuePrompt = ui.prompt(
    "EDIT " + headers[column - 1],
    "Nilai saat ini:\n" + oldValue +
    "\n\nMasukkan nilai pengganti:",
    ui.ButtonSet.OK_CANCEL
  );

  if (valuePrompt.getSelectedButton() !== ui.Button.OK) return;

  let newValue = valuePrompt.getResponseText();

  // NORMALISASI DATA TERTENTU
  if (column === 2) {
    newValue = String(newValue).trim().toUpperCase();
  }

  if (column === 7 || column === 8) {
    newValue = String(newValue).trim().toUpperCase();
  }

  if (column === 18 && String(newValue).trim() !== "") {
    newValue = Number(newValue);
  }

  // MENYIAPKAN DATA BARU TANPA MENGUBAH SHEET
  const proposedRow = [...currentRow];

  proposedRow[column - 1] = newValue;

  // MEMERIKSA ID DUPLIKAT
  const duplicateId = data.some((row, index) =>
    index !== matchingIndexes[0] &&
    String(row[0]).trim() === id
  );

  if (duplicateId) {
    ui.alert("ID_SOAL duplikat. Perubahan dibatalkan.");
    return;
  }

  // VALIDASI SATU BARIS
  const errors = validasiSatuBarisSoal_(proposedRow);

  if (errors.length > 0) {

    ui.alert(
      "PERUBAHAN DITOLAK\n\n" +
      "Data baru tidak memenuhi aturan BANK_SOAL.\n\n" +
      errors.join("\n") +
      "\n\nNilai lama tetap dipertahankan."
    );

    return;
  }

  // MENYIMPAN JIKA VALID
  sheet.getRange(
    rowNumber, column
  ).setValue(newValue);

  ui.alert(
    "EDIT SOAL BERHASIL\n\n" +
    "ID_SOAL: " + id +
    "\nKolom: " + headers[column - 1] +
    "\n\nData berhasil diperbarui dan lolos validasi."
  );

}


// ==================================================
// VALIDASI SATU BARIS SOAL
// ==================================================

function validasiSatuBarisSoal_(row) {

  const errors = [];

  const id = String(row[0] || "").trim();
  const type = String(row[1] || "").trim().toUpperCase();

  const addError = message => {
    errors.push(message);
  };

  // MEMERIKSA ID
  if (!id) {
    addError("ID_SOAL kosong.");
  }

  // MEMERIKSA JENIS SOAL
  const validTypes = ["PG", "PGK", "MJ", "IS", "UR"];

  if (!validTypes.includes(type)) {
    addError("JENIS_SOAL harus PG, PGK, MJ, IS, atau UR.");
  }

  // MEMERIKSA DATA WAJIB
  const requiredFields = [
    [2, "MATA PELAJARAN"],
    [3, "KELAS"],
    [4, "MATERI"],
    [5, "TUJUAN PEMBELAJARAN"],
    [6, "LEVEL BLOOM"],
    [7, "LEVEL KOGNITIF"],
    [8, "PERTANYAAN"]
  ];

  requiredFields.forEach(([column, label]) => {
    if (!String(row[column] || "").trim()) {
      addError(label + " kosong.");
    }
  });

  // MEMERIKSA LEVEL BLOOM
  const bloom = String(row[6] || "").trim();

  if (bloom && !/^C[1-6]$/i.test(bloom)) {
    addError("LEVEL_BLOOM harus C1 sampai C6.");
  }

  // MEMERIKSA LEVEL KOGNITIF
  const cognitive = String(row[7] || "").trim();

  if (cognitive && !/^L[1-3]$/i.test(cognitive)) {
    addError("LEVEL_KOGNITIF harus L1 sampai L3.");
  }

  // MEMERIKSA POIN
  const points = row[17];

  const validPoints =
    String(points).trim() !== "" &&
    Number.isInteger(Number(points)) &&
    Number(points) >= 0;

  if (!validPoints) {
    addError("POIN harus berupa bilangan bulat minimal 0.");
  }

  // MEMERIKSA PG DAN PGK
  if (type === "PG" || type === "PGK") {

    const options = row.slice(9, 13).map(value =>
      String(value || "").trim()
    );

    const key = String(row[13] || "")
      .split(",")
      .map(value => value.trim().toUpperCase())
      .filter(Boolean);

    if (options.some(option => !option)) {
      addError("OPSI A sampai D harus lengkap.");
    }

    const filledOptions = options.filter(Boolean);

    if (new Set(filledOptions).size !== filledOptions.length) {
      addError("Isi setiap opsi jawaban harus berbeda.");
    }

    if (
      key.some(value =>
        !["A", "B", "C", "D"].includes(value)
      ) ||
      new Set(key).size !== key.length ||
      (type === "PG" && key.length !== 1) ||
      (type === "PGK" && key.length < 2)
    ) {
      addError("KUNCI_JAWABAN tidak valid.");
    }
  }

  // MEMERIKSA MENJODOHKAN
  if (type === "MJ") {

    const rawPairs = String(row[15] || "").trim();

    if (!rawPairs) {

      addError("PASANGAN kosong.");

    } else {

      const pairs = rawPairs.split(";").map(pair => {

        const parts = pair.split("=");

        if (parts.length < 2) {
          return null;
        }

        return {
          left: parts[0].trim(),
          right: parts.slice(1).join("=").trim()
        };

      });

      if (pairs.some(pair =>
        !pair || !pair.left || !pair.right
      )) {

        addError(
          "Format PASANGAN harus kiri=kanan; kiri=kanan."
        );

      } else {

        const leftOptions = pairs.map(pair => pair.left);
        const rightOptions = pairs.map(pair => pair.right);

        if (new Set(leftOptions).size !== pairs.length) {
          addError("Bagian kiri pasangan tidak boleh sama.");
        }

        if (new Set(rightOptions).size !== pairs.length) {
          addError("Jawaban kanan pasangan tidak boleh sama.");
        }

        if (
          validPoints &&
          Number(points) < pairs.length
        ) {
          addError(
            "POIN harus minimal sama dengan jumlah pasangan."
          );
        }

      }
    }
  }

  // MEMERIKSA ISIAN SINGKAT
  if (
    type === "IS" &&
    !String(row[14] || "").trim()
  ) {
    addError("JAWABAN_DITERIMA kosong.");
  }

  // MEMERIKSA URAIAN
  if (
    type === "UR" &&
    !String(row[16] || "").trim()
  ) {
    addError("RUBRIK kosong.");
  }

  return errors;

}

// ==================================================
// 9. HAPUS SOAL
// ==================================================


 // ==================================================
 // 9. HAPUS SOAL
 // ==================================================

function hapusSoal() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANK_SOAL");

  if (!sheet) {
    ui.alert("Sheet BANK_SOAL tidak ditemukan.");
    return;
  }

  if (sheet.getLastRow() < 2) {
    ui.alert("BANK_SOAL belum memiliki data soal.");
    return;
  }

  // MEMINTA ID SOAL
  const idPrompt = ui.prompt(
    "HAPUS SOAL",
    "Masukkan ID_SOAL yang akan dihapus:",
    ui.ButtonSet.OK_CANCEL
  );

  if (idPrompt.getSelectedButton() !== ui.Button.OK) return;

  const id = idPrompt.getResponseText().trim();

  if (!id) {
    ui.alert("ID_SOAL tidak boleh kosong.");
    return;
  }

  // MENCARI ID SOAL
  const lastRow = sheet.getLastRow();

  const data = sheet.getRange(
    2, 1, lastRow - 1, 21
  ).getValues();

  const matchingIndexes = [];

  data.forEach((row, index) => {
    if (String(row[0]).trim() === id) {
      matchingIndexes.push(index);
    }
  });

  if (matchingIndexes.length === 0) {
    ui.alert("ID_SOAL tidak ditemukan.");
    return;
  }

  // MENCEGAH PENGHAPUSAN ID DUPLIKAT
  if (matchingIndexes.length > 1) {
    ui.alert(
      "ID_SOAL DUPLIKAT\n\n" +
      "Ditemukan lebih dari satu soal dengan ID " + id +
      ". Perbaiki ID terlebih dahulu sebelum menghapus."
    );
    return;
  }

  const rowNumber = matchingIndexes[0] + 2;

  const question = String(data[matchingIndexes[0]][8] || "");

  const type = String(data[matchingIndexes[0]][1] || "");

  // MEMERIKSA RIWAYAT PENGGUNAAN SOAL
  const mappingSheet = ss.getSheetByName(
    "PEMETAAN_SOAL_UJIAN"
  );

  const logSheet = ss.getSheetByName("LOG_SOAL_FORM");

  let usedInExam = false;
  let usedInLog = false;

  if (mappingSheet && mappingSheet.getLastRow() >= 2) {

    const mappingData = mappingSheet.getRange(
      2, 1, mappingSheet.getLastRow() - 1, 15
    ).getValues();

    usedInExam = mappingData.some(row =>
      String(row[1]).trim() === id
    );

  }

  if (logSheet && logSheet.getLastRow() >= 2) {

    const logData = logSheet.getRange(
      2, 1, logSheet.getLastRow() - 1, 3
    ).getValues();

    usedInLog = logData.some(row =>
      String(row[1]).trim() === id
    );

  }

  // MENYUSUN INFORMASI KONFIRMASI
  let warning = "";

  if (usedInExam || usedInLog) {
    warning =
      "\n\nPERINGATAN:\n" +
      "Soal ini memiliki riwayat penggunaan dalam ujian.\n" +
      "Data pemetaan dan log ujian tidak akan ikut dihapus.";
  }

  // KONFIRMASI PENGHAPUSAN
  const confirmation = ui.alert(
    "KONFIRMASI HAPUS SOAL",
    "ID_SOAL: " + id +
    "\nJENIS SOAL: " + type +
    "\n\nPERTANYAAN:\n" + question +
    warning +
    "\n\nSoal akan dihapus permanen dari BANK_SOAL." +
    "\n\nYakin ingin melanjutkan?",
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    ui.alert("Penghapusan dibatalkan.");
    return;
  }

  // MENGHAPUS BARIS SOAL
  sheet.deleteRow(rowNumber);

  ui.alert(
    "PENGHAPUSAN BERHASIL\n\n" +
    "ID_SOAL: " + id +
    "\nSoal berhasil dihapus dari BANK_SOAL." +
    "\nRiwayat pemetaan dan penggunaan ujian tetap dipertahankan."
  );

}


// ==================================================
// 10. DETEKSI PERTANYAAN DUPLIKAT
// ==================================================


function deteksiDuplikatSoal() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANK_SOAL");

  if (!sheet) {
    ui.alert("Sheet BANK_SOAL tidak ditemukan.");
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert("BANK_SOAL belum memiliki data soal.");
    return;
  }

  const data = sheet.getRange(
    2, 1, lastRow - 1, 21
  ).getValues();

  const groups = new Map();

  let checkedCount = 0;
  let emptyCount = 0;

  data.forEach((row, index) => {

    const rowNumber = index + 2;
    const id = String(row[0] || "").trim();
    const question = String(row[8] || "").trim();

    if (!question) {
      emptyCount++;
      return;
    }

    checkedCount++;

    // Normalisasi teks pertanyaan
    const normalized = question
      .toLowerCase()
      .normalize("NFKC")
      .replace(/[^\p{L}\p{N}]+/gu, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (!normalized) return;

    if (!groups.has(normalized)) {
      groups.set(normalized, []);
    }

    groups.get(normalized).push({
      id: id || "TANPA ID",
      row: rowNumber,
      question: question
    });

  });

  const duplicates = [];

  let duplicateQuestionCount = 0;

  groups.forEach(items => {

    if (items.length > 1) {

      duplicates.push(items);
      duplicateQuestionCount += items.length;

    }

  });

  let message =
    "HASIL DETEKSI PERTANYAAN DUPLIKAT\n\n" +
    "Jumlah pertanyaan diperiksa: " + checkedCount +
    "\nPertanyaan kosong dilewati: " + emptyCount +
    "\nKelompok duplikat: " + duplicates.length +
    "\nTotal soal terindikasi duplikat: " +
    duplicateQuestionCount;

  if (duplicates.length > 0) {

    message += "\n\nRINCIAN DUPLIKAT:\n";

    duplicates.forEach((items, index) => {

      message += "\nKELOMPOK " + (index + 1) + "\n";

      items.forEach(item => {

        message +=
          "Baris " + item.row +
          " | ID: " + item.id + "\n";

      });

      message +=
        "Pertanyaan: " + items[0].question + "\n";

    });

    ui.alert(
      "DUPLIKAT DITEMUKAN",
      message,
      ui.ButtonSet.OK
    );

  } else {

    ui.alert(
      "TIDAK ADA DUPLIKAT",
      message +
      "\n\nTidak ditemukan pertanyaan dengan teks yang sama setelah normalisasi.",
      ui.ButtonSet.OK
    );

  }

}


// ==================================================
// 11. PENCARIAN DAN FILTER BANK SOAL
// ==================================================


function cariFilterBankSoal() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANK_SOAL");

  if (!sheet) {
    ui.alert("Sheet BANK_SOAL tidak ditemukan.");
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert("BANK_SOAL belum memiliki data soal.");
    return;
  }

  const filters = [
    ["ID_SOAL", "Masukkan ID soal:"],
    ["MATA PELAJARAN", "Masukkan mata pelajaran:"],
    ["KELAS", "Masukkan kelas:"],
    ["MATERI", "Masukkan kata kunci materi:"],
    ["JENIS SOAL", "Masukkan jenis soal (PG/PGK/MJ/IS/UR):"],
    ["LEVEL BLOOM", "Masukkan level Bloom (C1-C6):"],
    ["LEVEL KOGNITIF", "Masukkan level kognitif (L1-L3):"],
    ["PERTANYAAN", "Masukkan kata kunci pertanyaan:"]
  ];

  const criteria = {};

  for (const filter of filters) {

    const response = ui.prompt(
      "PENCARIAN BANK SOAL",
      filter[1] + "\n\nKosongkan jika kriteria ini tidak digunakan.",
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Pencarian dibatalkan.");
      return;
    }

    criteria[filter[0]] = String(
      response.getResponseText() || ""
    ).trim().toLowerCase();

  }

  // Validasi jenis soal jika diisi
  const validTypes = ["pg", "pgk", "mj", "is", "ur"];

  if (
    criteria["JENIS SOAL"] &&
    !validTypes.includes(criteria["JENIS SOAL"].toLowerCase())
  ) {
    ui.alert(
      "JENIS SOAL TIDAK VALID",
      "Gunakan salah satu jenis berikut: PG, PGK, MJ, IS, UR.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Validasi level Bloom jika diisi
  const validBloom = ["c1", "c2", "c3", "c4", "c5", "c6"];

  if (
    criteria["LEVEL BLOOM"] &&
    !validBloom.includes(criteria["LEVEL BLOOM"])
  ) {
    ui.alert(
      "LEVEL BLOOM TIDAK VALID",
      "Gunakan level C1 sampai C6.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Validasi level kognitif jika diisi
  const validCognitive = ["l1", "l2", "l3"];

  if (
    criteria["LEVEL KOGNITIF"] &&
    !validCognitive.includes(criteria["LEVEL KOGNITIF"])
  ) {
    ui.alert(
      "LEVEL KOGNITIF TIDAK VALID",
      "Gunakan level L1, L2, atau L3.",
      ui.ButtonSet.OK
    );
    return;
  }

  const data = sheet.getRange(
    2, 1, lastRow - 1, 21
  ).getValues();

  const columnMap = {
    "ID_SOAL": 0,
    "MATA PELAJARAN": 2,
    "KELAS": 3,
    "MATERI": 4,
    "JENIS SOAL": 1,
    "LEVEL BLOOM": 6,
    "LEVEL KOGNITIF": 7,
    "PERTANYAAN": 8
  };

  const activeCriteria = filters.filter(
    filter => criteria[filter[0]] !== ""
  );

  if (activeCriteria.length === 0) {
    ui.alert(
      "KRITERIA KOSONG",
      "Masukkan minimal satu kriteria pencarian.",
      ui.ButtonSet.OK
    );
    return;
  }

  const matches = [];

  data.forEach((row, index) => {

    // Lewati baris tanpa ID dan pertanyaan
    const id = String(row[0] || "").trim();
    const question = String(row[8] || "").trim();

    if (!id && !question) {
      return;
    }

    const isMatch = activeCriteria.every(filter => {

      const field = filter[0];
      const keyword = criteria[field];

      const value = String(
        row[columnMap[field]] || ""
      ).trim().toLowerCase();

      return value.includes(keyword);

    });

    if (isMatch) {
      matches.push(index + 2);
    }

  });

  // Tampilkan seluruh baris terlebih dahulu
  sheet.showRows(1, sheet.getMaxRows());

  if (matches.length === 0) {

    ui.alert(
      "HASIL PENCARIAN",
      "Tidak ditemukan soal yang sesuai dengan kriteria.",
      ui.ButtonSet.OK
    );

    return;
  }

  // Menyembunyikan baris yang tidak sesuai
  const matchSet = new Set(matches);

  let startHidden = null;

  for (
    let row = 2;
    row <= lastRow + 1;
    row++
  ) {

    const shouldHide =
      row <= lastRow && !matchSet.has(row);

    if (shouldHide && startHidden === null) {
      startHidden = row;
    }

    if (!shouldHide && startHidden !== null) {

      sheet.hideRows(
        startHidden,
        row - startHidden
      );

      startHidden = null;
    }

  }

  ui.alert(
    "PENCARIAN SELESAI",
    "Jumlah kriteria digunakan: " + activeCriteria.length +
    "\nJumlah soal ditemukan: " + matches.length +
    "\n\nBaris yang tidak sesuai telah disembunyikan." +
    "\n\nGunakan menu Reset Filter Bank Soal untuk menampilkan semua soal kembali.",
    ui.ButtonSet.OK
  );

}

// ==================================================
// 12. RESET FILTER BANK SOAL
// ==================================================

function resetFilterBankSoal() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BANK_SOAL");

  if (!sheet) {
    ui.alert("Sheet BANK_SOAL tidak ditemukan.");
    return;
  }

  // Menghapus filter jika ada
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  // Menampilkan kembali semua baris
  sheet.showRows(1, sheet.getMaxRows());

  ui.alert(
    "RESET BERHASIL",
    "Semua baris pada BANK_SOAL sekarang ditampilkan kembali.",
    ui.ButtonSet.OK
  );

}

function siapkanSheetAnalisis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const struktur = {
    DAFTAR_UJIAN: [
      "ID_UJIAN", "NAMA_UJIAN", "FORM_ID", "FORM_URL",
      "RESPONSE_SPREADSHEET_ID", "RESPONSE_SHEET_NAME",
      "TANGGAL_DIBUAT", "TANGGAL_ANALISIS", "JUMLAH_SOAL", "STATUS"
    ],
    PEMETAAN_SOAL_UJIAN: [
      "ID_UJIAN", "ID_SOAL", "ITEM_ID_FORM", "NOMOR_ITEM",
      "SUB_ITEM_KE", "JENIS_SOAL", "TEKS_PERTANYAAN",
      "POIN_MAKSIMAL", "KUNCI_JAWABAN", "JAWABAN_DITERIMA",
      "RUBRIK", "MATERI", "LEVEL_BLOOM",
      "LEVEL_KOGNITIF", "STATUS_PENILAIAN"
    ],
    ANALISIS_UJIAN: [
      "ID_UJIAN", "NAMA_UJIAN", "TANGGAL_ANALISIS",
      "JUMLAH_PESERTA", "JUMLAH_DINILAI",
      "JUMLAH_BELUM_DINILAI", "NILAI_RATA_RATA",
      "NILAI_TERTINGGI", "NILAI_TERENDAH", "KKM",
      "JUMLAH_TUNTAS", "JUMLAH_TIDAK_TUNTAS",
      "PERSENTASE_TUNTAS"
    ],
    ANALISIS_PESERTA: [
      "ID_UJIAN", "ID_RESPONDEN", "NAMA", "KELAS",
      "EMAIL", "NILAI", "STATUS_KETUNTASAN",
      "JUMLAH_BENAR", "JUMLAH_SALAH",
      "JUMLAH_BELUM_DINILAI", "WAKTU_SUBMIT"
    ],
    ANALISIS_PER_SOAL: [
      "ID_UJIAN", "ID_SOAL", "ITEM_ID_FORM",
      "NOMOR_SOAL", "MATERI", "JENIS_SOAL",
      "POIN_MAKSIMAL", "JUMLAH_RESPONDEN",
      "JUMLAH_BENAR", "JUMLAH_SALAH",
      "JUMLAH_BELUM_DINILAI", "TINGKAT_KESUKARAN",
      "KATEGORI_KESUKARAN", "DAYA_PEMBEDA",
      "KATEGORI_DAYA_PEMBEDA"
    ]
  };

  for (const namaSheet in struktur) {
    let sheet = ss.getSheetByName(namaSheet);

    if (!sheet) {
      sheet = ss.insertSheet(namaSheet);
    }

    if (sheet.getLastRow() === 0) {
      const header = struktur[namaSheet];

      sheet.getRange(1, 1, 1, header.length)
        .setValues([header])
        .setFontWeight("bold")
        .setBackground("#d9ead3");

      sheet.setFrozenRows(1);
    }
  }

  const dashboard = ss.getSheetByName("DASHBOARD_UJIAN");

  if (dashboard && dashboard.getLastRow() === 0) {
    dashboard.getRange("A1").setValue("DASHBOARD ANALISIS UJIAN");
  }

  SpreadsheetApp.getUi().alert(
    "Struktur sheet analisis berhasil disiapkan."
  );
}


function siapkanDaftarUjian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DAFTAR_UJIAN");

  const header = [
    "ID_UJIAN",
    "NAMA_UJIAN",
    "FORM_ID",
    "FORM_URL",
    "RESPONSE_SPREADSHEET_ID",
    "RESPONSE_SHEET_NAME",
    "TANGGAL_DIBUAT",
    "TANGGAL_ANALISIS",
    "JUMLAH_SOAL",
    "STATUS"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert("DAFTAR_UJIAN berhasil disiapkan.");
}

function siapkanPemetaanSoalUjian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");

  const header = [
    "ID_UJIAN",
    "ID_SOAL",
    "ITEM_ID_FORM",
    "NOMOR_ITEM",
    "SUB_ITEM_KE",
    "JENIS_SOAL",
    "TEKS_PERTANYAAN",
    "POIN_MAKSIMAL",
    "KUNCI_JAWABAN",
    "JAWABAN_DITERIMA",
    "RUBRIK",
    "MATERI",
    "LEVEL_BLOOM",
    "LEVEL_KOGNITIF",
    "STATUS_PENILAIAN"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}


function siapkanAnalisisUjian() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ANALISIS_UJIAN");

  const header = [
    "ID_UJIAN",
    "NAMA_UJIAN",
    "TANGGAL_ANALISIS",
    "JUMLAH_PESERTA",
    "JUMLAH_DINILAI",
    "JUMLAH_BELUM_DINILAI",
    "NILAI_RATA_RATA",
    "NILAI_TERTINGGI",
    "NILAI_TERENDAH",
    "KKM",
    "JUMLAH_TUNTAS",
    "JUMLAH_TIDAK_TUNTAS",
    "PERSENTASE_TUNTAS"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}

function siapkanAnalisisPeserta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ANALISIS_PESERTA");

  const header = [
    "ID_UJIAN",
    "ID_RESPONDEN",
    "NAMA",
    "KELAS",
    "EMAIL",
    "NILAI",
    "STATUS_KETUNTASAN",
    "JUMLAH_BENAR",
    "JUMLAH_SALAH",
    "JUMLAH_BELUM_DINILAI",
    "WAKTU_SUBMIT"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}

function siapkanAnalisisPerSoal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ANALISIS_PER_SOAL");

  const header = [
    "ID_UJIAN",
    "ID_SOAL",
    "ITEM_ID_FORM",
    "NOMOR_SOAL",
    "MATERI",
    "JENIS_SOAL",
    "POIN_MAKSIMAL",
    "JUMLAH_RESPONDEN",
    "JUMLAH_BENAR",
    "JUMLAH_SALAH",
    "JUMLAH_BELUM_DINILAI",
    "TINGKAT_KESUKARAN",
    "KATEGORI_KESUKARAN",
    "DAYA_PEMBEDA",
    "KATEGORI_DAYA_PEMBEDA"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
  }
}

function analisisHasilUjian() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const input = ui.prompt(
    "ANALISIS HASIL UJIAN",
    "Masukkan ID_UJIAN yang ingin dianalisis:",
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const idUjian = input.getResponseText().trim();

  if (!idUjian) {
    ui.alert("ID_UJIAN tidak boleh kosong.");
    return;
  }

  try {
    const daftarSheet = ss.getSheetByName("DAFTAR_UJIAN");
    const mappingSheet = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");
    const hasilSheet = ss.getSheetByName("ANALISIS_PESERTA");
    const pengaturanSheet = ss.getSheetByName("PENGATURAN");

    if (!daftarSheet || !mappingSheet || !hasilSheet || !pengaturanSheet) {
      throw new Error("Salah satu sheet analisis tidak ditemukan.");
    }

    const daftarData = daftarSheet.getDataRange().getValues();

    const ujian = daftarData.find((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (!ujian) throw new Error("ID_UJIAN tidak ditemukan.");

    const pengaturan = pengaturanSheet.getDataRange().getValues();

    const barisKKM = pengaturan.find(row =>
      String(row[0]).trim() === "KKM"
    );

    if (!barisKKM || isNaN(Number(barisKKM[1]))) {
      throw new Error("Pengaturan KKM tidak valid.");
    }

    const kkm = Number(barisKKM[1]);

    const mappingData = mappingSheet.getDataRange().getValues();

    const mappings = mappingData.filter((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (mappings.length === 0) {
      throw new Error("Pemetaan soal tidak ditemukan.");
    }

    const totalPoin = mappings.reduce(
      (sum, row) => sum + Number(row[7] || 0),
      0
    );

    if (totalPoin <= 0) {
      throw new Error("Total poin maksimal tidak valid.");
    }

    const form = FormApp.openById(String(ujian[2]));
    const responses = form.getResponses();

    if (responses.length === 0) {
      throw new Error("Belum ada respons peserta.");
    }

    const hasilBaru = [];

    responses.forEach(response => {
      const itemResponses = response.getItemResponses();
      const skorResponses = response.getGradableItemResponses();

      let nama = "";
      let kelas = "";

      itemResponses.forEach(itemResponse => {
        const judul = itemResponse.getItem().getTitle().trim();
        const jawaban = itemResponse.getResponse();

        if (judul === "Nama Lengkap") {
          nama = String(jawaban || "");
        }

        if (judul === "Kelas") {
          kelas = String(jawaban || "");
        }
      });

      const skorPerItem = {};

      skorResponses.forEach(itemResponse => {
        const itemId = String(itemResponse.getItem().getId());
        skorPerItem[itemId] = itemResponse.getScore();
      });

      let skorDiperoleh = 0;
      let jumlahBenar = 0;
      let jumlahSalah = 0;
      let jumlahBelumDinilai = 0;

      mappings.forEach(mapping => {
        const itemId = String(mapping[2]);
        const poinMaksimal = Number(mapping[7] || 0);
        const skor = skorPerItem[itemId];

        if (
          skor === undefined ||
          skor === null ||
          !Number.isFinite(Number(skor))
        ) {
          jumlahBelumDinilai++;
          return;
        }

        const skorItem = Number(skor);
        skorDiperoleh += skorItem;

        if (skorItem >= poinMaksimal) {
          jumlahBenar++;
        } else {
          jumlahSalah++;
        }
      });

      let nilai = "";
      let statusKetuntasan = "BELUM_SELESAI";

      if (jumlahBelumDinilai === 0) {
        nilai = Math.round(
          (skorDiperoleh / totalPoin) * 10000
        ) / 100;

        statusKetuntasan = nilai >= kkm
          ? "TUNTAS"
          : "TIDAK_TUNTAS";
      }

      const email = response.getRespondentEmail() || "";

      hasilBaru.push([
        idUjian,
        response.getId() || String(response.getTimestamp().getTime()),
        nama,
        kelas,
        email,
        nilai,
        statusKetuntasan,
        jumlahBenar,
        jumlahSalah,
        jumlahBelumDinilai,
        response.getTimestamp()
      ]);
    });

    const lastRow = hasilSheet.getLastRow();
    let dataLama = [];

    if (lastRow > 1) {
      dataLama = hasilSheet.getRange(
        2, 1, lastRow - 1, 11
      ).getValues();
    }

    const dataGabungan = dataLama.filter(row =>
      String(row[0]).trim() !== idUjian
    ).concat(hasilBaru);

    if (lastRow > 1) {
      hasilSheet.getRange(
        2, 1, lastRow - 1, 11
      ).clearContent();
    }

    if (dataGabungan.length > 0) {
      hasilSheet.getRange(
        2, 1, dataGabungan.length, 11
      ).setValues(dataGabungan);
    }

    ui.alert(
      "ANALISIS PESERTA BERHASIL",
      "Nama ujian: " + ujian[1] +
      "\nJumlah peserta: " + hasilBaru.length +
      "\nTotal poin maksimal: " + totalPoin +
      "\nKKM: " + kkm +
      "\nHasil tersimpan di ANALISIS_PESERTA.",
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert("ANALISIS GAGAL:\n" + error.message);
  }
}

function analisisPerSoal() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const input = ui.prompt(
    "ANALISIS PER SOAL",
    "Masukkan ID_UJIAN yang ingin dianalisis:",
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const idUjian = input.getResponseText().trim();

  if (!idUjian) {
    ui.alert("ID_UJIAN tidak boleh kosong.");
    return;
  }

  try {
    const daftarSheet = ss.getSheetByName("DAFTAR_UJIAN");
    const mappingSheet = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");
    const hasilSheet = ss.getSheetByName("ANALISIS_PER_SOAL");

    if (!daftarSheet || !mappingSheet || !hasilSheet) {
      throw new Error("Sheet analisis tidak ditemukan.");
    }

    // Mencari data ujian
    const daftarData = daftarSheet.getDataRange().getValues();

    const ujian = daftarData.find((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (!ujian) {
      throw new Error("ID_UJIAN tidak ditemukan.");
    }

    // Membaca pemetaan soal
    const mappingData = mappingSheet.getDataRange().getValues();

    const mappings = mappingData.filter((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (mappings.length === 0) {
      throw new Error("Pemetaan soal tidak ditemukan.");
    }

    // Membaca respons Google Form
    const form = FormApp.openById(String(ujian[2]));
    const responses = form.getResponses();

    if (responses.length === 0) {
      throw new Error("Belum ada respons peserta.");
    }

    // Mengambil skor setiap peserta berdasarkan ID item
    const peserta = responses.map(response => {

      const skorPerItem = {};

      response.getGradableItemResponses().forEach(itemResponse => {
        const itemId = String(itemResponse.getItem().getId());
        skorPerItem[itemId] = itemResponse.getScore();
      });

      let totalSkor = 0;
      let semuaDinilai = true;

      mappings.forEach(mapping => {
        const itemId = String(mapping[2]);
        const skor = skorPerItem[itemId];

        if (skor === undefined || skor === null ||
            !Number.isFinite(Number(skor))) {
          semuaDinilai = false;
        } else {
          totalSkor += Number(skor);
        }
      });

      return {
        skorPerItem: skorPerItem,
        totalSkor: totalSkor,
        semuaDinilai: semuaDinilai
      };
    });

    // Kelompok peserta untuk daya pembeda
    const pesertaLengkap = peserta
      .filter(p => p.semuaDinilai)
      .sort((a, b) => b.totalSkor - a.totalSkor);

    const jumlahKelompok = Math.max(
      1,
      Math.floor(pesertaLengkap.length * 0.27)
    );

    const kelompokAtas = pesertaLengkap.slice(0, jumlahKelompok);

    const kelompokBawah = pesertaLengkap.slice(
      pesertaLengkap.length - jumlahKelompok
    );

    const hasilBaru = [];

    // Menghitung setiap soal
    mappings.forEach(mapping => {

      const idSoal = mapping[1];
      const itemId = String(mapping[2]);
      const nomorSoal = mapping[3];
      const jenisSoal = mapping[5];
      const poinMaksimal = Number(mapping[7] || 0);
      const materi = mapping[11];

      let jumlahBenar = 0;
      let jumlahSalah = 0;
      let jumlahBelumDinilai = 0;
      let totalSkorItem = 0;
      let jumlahDinilai = 0;

      peserta.forEach(p => {

        const skor = p.skorPerItem[itemId];

        if (skor === undefined || skor === null ||
            !Number.isFinite(Number(skor))) {
          jumlahBelumDinilai++;
          return;
        }

        const skorItem = Number(skor);
                
        jumlahDinilai++;
        totalSkorItem += skorItem;

        if (skorItem >= poinMaksimal) {
          jumlahBenar++;
        } else {
          jumlahSalah++;
        }
      });

      // Indeks kesukaran
      let tingkatKesukaran = "";
      let kategoriKesukaran = "";

      if (jumlahDinilai > 0 && poinMaksimal > 0) {

        tingkatKesukaran =
          totalSkorItem / (jumlahDinilai * poinMaksimal);

        if (tingkatKesukaran <= 0.30) {
          kategoriKesukaran = "SULIT";
        } else if (tingkatKesukaran <= 0.70) {
          kategoriKesukaran = "SEDANG";
        } else {
          kategoriKesukaran = "MUDAH";
        }
      }

      // Indeks daya pembeda
      let dayaPembeda = "";
      let kategoriDayaPembeda = "";

      if (
        pesertaLengkap.length >= 2 &&
        kelompokAtas.length > 0 &&
        kelompokBawah.length > 0 &&
        poinMaksimal > 0
      ) {

        function rataRataKelompok(kelompok) {
          const total = kelompok.reduce((sum, p) => {
            return sum +
              Number(p.skorPerItem[itemId] || 0) / poinMaksimal;
          }, 0);

          return total / kelompok.length;
        }

        dayaPembeda =
          rataRataKelompok(kelompokAtas) -
          rataRataKelompok(kelompokBawah);

        if (dayaPembeda >= 0.40) {
          kategoriDayaPembeda = "SANGAT BAIK";
        } else if (dayaPembeda >= 0.30) {
          kategoriDayaPembeda = "BAIK";
        } else if (dayaPembeda >= 0.20) {
          kategoriDayaPembeda = "CUKUP";
        } else if (dayaPembeda >= 0) {
          kategoriDayaPembeda = "KURANG";
        } else {
          kategoriDayaPembeda = "BURUK";
        }
      }

      hasilBaru.push([
        idUjian,
        idSoal,
        itemId,
        nomorSoal,
        materi,
        jenisSoal,
        poinMaksimal,
        peserta.length,
        jumlahBenar,
        jumlahSalah,
        jumlahBelumDinilai,
        tingkatKesukaran,
        kategoriKesukaran,
        dayaPembeda,
        kategoriDayaPembeda
      ]);
    });

    // Mempertahankan hasil ujian lain
    const lastRow = hasilSheet.getLastRow();

    let dataLama = [];

    if (lastRow > 1) {
      dataLama = hasilSheet.getRange(
        2, 1, lastRow - 1, 15
      ).getValues();
    }

    const dataGabungan = dataLama.filter(row =>
      String(row[0]).trim() !== idUjian
    ).concat(hasilBaru);

    if (lastRow > 1) {
      hasilSheet.getRange(
        2, 1, lastRow - 1, 15
      ).clearContent();
    }

    if (dataGabungan.length > 0) {
      hasilSheet.getRange(
        2, 1, dataGabungan.length, 15
      ).setValues(dataGabungan);
    }

    // Format angka indeks
    const barisHasil = hasilSheet.getLastRow();

    if (barisHasil > 1) {
      hasilSheet.getRange(
        2, 12, barisHasil - 1, 1
      ).setNumberFormat("0.00");

      hasilSheet.getRange(
        2, 14, barisHasil - 1, 1
      ).setNumberFormat("0.00");
    }

    ui.alert(
      "ANALISIS PER SOAL BERHASIL",
      "Nama ujian: " + ujian[1] +
      "\nJumlah soal: " + hasilBaru.length +
      "\nJumlah peserta: " + peserta.length +
      "\nPeserta yang seluruh soalnya sudah dinilai: " +
      pesertaLengkap.length +
      "\nHasil tersimpan di ANALISIS_PER_SOAL.",
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert("ANALISIS PER SOAL GAGAL:\n" + error.message);
  }
}


function analisisRingkasanUjian() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const input = ui.prompt(
    "RINGKASAN ANALISIS UJIAN",
    "Masukkan ID_UJIAN yang ingin dianalisis:",
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const idUjian = input.getResponseText().trim();

  if (!idUjian) {
    ui.alert("ID_UJIAN tidak boleh kosong.");
    return;
  }

  try {
    const daftarSheet = ss.getSheetByName("DAFTAR_UJIAN");
    const pesertaSheet = ss.getSheetByName("ANALISIS_PESERTA");
    const hasilSheet = ss.getSheetByName("ANALISIS_UJIAN");
    const pengaturanSheet = ss.getSheetByName("PENGATURAN");

    if (!daftarSheet || !pesertaSheet || !hasilSheet || !pengaturanSheet) {
      throw new Error("Sheet analisis tidak ditemukan.");
    }

    // Mencari data ujian
    const daftarData = daftarSheet.getDataRange().getValues();

    const ujian = daftarData.find((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (!ujian) {
      throw new Error("ID_UJIAN tidak ditemukan.");
    }

    // Membaca KKM
    const pengaturan = pengaturanSheet.getDataRange().getValues();

    const barisKKM = pengaturan.find(row =>
      String(row[0]).trim() === "KKM"
    );

    if (!barisKKM || isNaN(Number(barisKKM[1]))) {
      throw new Error("KKM tidak valid.");
    }

    const kkm = Number(barisKKM[1]);

    // Membaca hasil peserta
    const pesertaData = pesertaSheet.getDataRange().getValues();

    const peserta = pesertaData.filter((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (peserta.length === 0) {
      throw new Error("Belum ada hasil analisis peserta.");
    }

    const jumlahPeserta = peserta.length;

    const pesertaDinilai = peserta.filter(row =>
      row[6] === "TUNTAS" || row[6] === "TIDAK_TUNTAS"
    );

    const pesertaBelumDinilai = jumlahPeserta - pesertaDinilai.length;

    const nilaiPeserta = pesertaDinilai.map(row =>
      Number(row[5])
    );

    const nilaiRataRata = nilaiPeserta.length > 0
      ? nilaiPeserta.reduce((sum, nilai) => sum + nilai, 0) /
        nilaiPeserta.length
      : "";

    const nilaiTertinggi = nilaiPeserta.length > 0
      ? Math.max(...nilaiPeserta)
      : "";

    const nilaiTerendah = nilaiPeserta.length > 0
      ? Math.min(...nilaiPeserta)
      : "";

    const jumlahTuntas = peserta.filter(row =>
      row[6] === "TUNTAS"
    ).length;

    const jumlahTidakTuntas = peserta.filter(row =>
      row[6] === "TIDAK_TUNTAS"
    ).length;

    const persentaseTuntas = pesertaDinilai.length > 0
      ? jumlahTuntas / pesertaDinilai.length * 100
      : 0;

    const hasilBaru = [
      idUjian,
      ujian[1],
      new Date(),
      jumlahPeserta,
      pesertaDinilai.length,
      pesertaBelumDinilai,
      nilaiRataRata,
      nilaiTertinggi,
      nilaiTerendah,
      kkm,
      jumlahTuntas,
      jumlahTidakTuntas,
      persentaseTuntas
    ];

    // Mempertahankan data ujian lain
    const lastRow = hasilSheet.getLastRow();

    let dataLama = [];

    if (lastRow > 1) {
      dataLama = hasilSheet.getRange(
        2, 1, lastRow - 1, 13
      ).getValues();
    }

    const dataGabungan = dataLama.filter(row =>
      String(row[0]).trim() !== idUjian
    );

    dataGabungan.push(hasilBaru);

    if (lastRow > 1) {
      hasilSheet.getRange(
        2, 1, lastRow - 1, 13
      ).clearContent();
    }

    hasilSheet.getRange(
      2, 1, dataGabungan.length, 13
    ).setValues(dataGabungan);

    // Format angka
    hasilSheet.getRange(
      2, 7, dataGabungan.length, 3
    ).setNumberFormat("0.00");

    hasilSheet.getRange(
      2, 13, dataGabungan.length, 1
    ).setNumberFormat("0.00");

    ui.alert(
      "RINGKASAN ANALISIS BERHASIL",
      "Nama ujian: " + ujian[1] +
      "\nJumlah peserta: " + jumlahPeserta +
      "\nSudah dinilai: " + pesertaDinilai.length +
      "\nBelum dinilai: " + pesertaBelumDinilai +
      "\nRata-rata nilai: " + nilaiRataRata.toFixed(2) +
      "\nNilai tertinggi: " + nilaiTertinggi +
      "\nNilai terendah: " + nilaiTerendah +
      "\nTuntas: " + jumlahTuntas +
      "\nTidak tuntas: " + jumlahTidakTuntas +
      "\nPersentase tuntas: " + persentaseTuntas.toFixed(2) + "%",
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert("RINGKASAN ANALISIS GAGAL:\n" + error.message);
  }
}


function buatDashboardUjian() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const input = ui.prompt(
    "DASHBOARD UJIAN",
    "Masukkan ID_UJIAN yang ingin ditampilkan:",
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const idUjian = input.getResponseText().trim();

  if (!idUjian) {
    ui.alert("ID_UJIAN tidak boleh kosong.");
    return;
  }

  try {
    const ringkasanSheet = ss.getSheetByName("ANALISIS_UJIAN");
    const pesertaSheet = ss.getSheetByName("ANALISIS_PESERTA");
    const soalSheet = ss.getSheetByName("ANALISIS_PER_SOAL");
    const dashboard = ss.getSheetByName("DASHBOARD_UJIAN");

    if (!ringkasanSheet || !pesertaSheet || !soalSheet || !dashboard) {
      throw new Error("Salah satu sheet dashboard tidak ditemukan.");
    }

    const ringkasanData = ringkasanSheet.getDataRange().getValues();

    const ringkasan = ringkasanData.find((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (!ringkasan) {
      throw new Error("Ringkasan ujian belum tersedia. Jalankan Ringkasan Analisis Ujian terlebih dahulu.");
    }

    const pesertaData = pesertaSheet.getDataRange().getValues();

    const peserta = pesertaData.filter((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    const soalData = soalSheet.getDataRange().getValues();

    const soal = soalData.filter((row, index) =>
      index > 0 && String(row[0]).trim() === idUjian
    );

    if (peserta.length === 0 || soal.length === 0) {
      throw new Error("Data analisis peserta atau soal belum tersedia.");
    }

    // Membersihkan dashboard lama
    dashboard.clear();
    dashboard.getCharts().forEach(chart => dashboard.removeChart(chart));

    // Judul dashboard
    dashboard.getRange("A1:H1").merge();
    dashboard.getRange("A1")
      .setValue("DASHBOARD ANALISIS UJIAN")
      .setFontSize(16)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    dashboard.getRange("A2").setValue("Nama Ujian");
    dashboard.getRange("B2").setValue(ringkasan[1]);

    dashboard.getRange("A3").setValue("ID Ujian");
    dashboard.getRange("B3").setValue(idUjian);

    // Ringkasan nilai
    dashboard.getRange("A5:B5").merge();
    dashboard.getRange("A5").setValue("RINGKASAN HASIL");

    const ringkasanTampilan = [
      ["Jumlah Peserta", ringkasan[3]],
      ["Sudah Dinilai", ringkasan[4]],
      ["Belum Dinilai", ringkasan[5]],
      ["Nilai Rata-rata", ringkasan[6]],
      ["Nilai Tertinggi", ringkasan[7]],
      ["Nilai Terendah", ringkasan[8]],
      ["KKM", ringkasan[9]],
      ["Jumlah Tuntas", ringkasan[10]],
      ["Jumlah Tidak Tuntas", ringkasan[11]],
      ["Persentase Tuntas (%)", ringkasan[12]]
    ];

    dashboard.getRange(6, 1, ringkasanTampilan.length, 2)
      .setValues(ringkasanTampilan);

    // Data grafik ketuntasan
    dashboard.getRange("D5:E5").setValues([
      ["Status", "Jumlah"]
    ]);

    dashboard.getRange("D6:E7").setValues([
      ["Tuntas", ringkasan[10]],
      ["Tidak Tuntas", ringkasan[11]]
    ]);

    const grafikKetuntasan = dashboard.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dashboard.getRange("D5:E7"))
      .setPosition(5, 4, 0, 0)
      .setOption("title", "Persentase Ketuntasan")
      .setOption("pieHole", 0.4)
      .setOption("width", 350)
      .setOption("height", 250)
      .build();

    dashboard.insertChart(grafikKetuntasan);

    // Data grafik distribusi nilai
    const kelompokNilai = [
      ["Rentang Nilai", "Jumlah Peserta"],
      ["0–59", 0],
      ["60–74", 0],
      ["75–84", 0],
      ["85–100", 0]
    ];

    peserta.forEach(row => {
      if (row[6] !== "TUNTAS" && row[6] !== "TIDAK_TUNTAS") return;

      const nilai = Number(row[5]);

      if (nilai < 60) kelompokNilai[1][1]++;
      else if (nilai < 75) kelompokNilai[2][1]++;
      else if (nilai < 85) kelompokNilai[3][1]++;
      else kelompokNilai[4][1]++;
    });

    dashboard.getRange(5, 7, kelompokNilai.length, 2)
      .setValues(kelompokNilai);

    const grafikNilai = dashboard.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dashboard.getRange("G5:H9"))
      .setPosition(20, 1, 0, 0)
      .setOption("title", "Distribusi Nilai Peserta")
      .setOption("legend", { position: "none" })
      .setOption("width", 350)
      .setOption("height", 250)
      .build();

    dashboard.insertChart(grafikNilai);

    // Tabel analisis per soal
    dashboard.getRange("A37").setValue("ANALISIS PER SOAL");

    const headerSoal = [
      "Nomor",
      "ID Soal",
      "Materi",
      "Kesukaran",
      "Kategori Kesukaran",
      "Daya Pembeda",
      "Kategori Daya Pembeda"
    ];

    dashboard.getRange(38, 1, 1, headerSoal.length)
      .setValues([headerSoal]);

    const tabelSoal = soal.map(row => [
      row[3],
      row[1],
      row[4],
      row[11],
      row[12],
      row[13],
      row[14]
    ]);

    dashboard.getRange(39, 1, tabelSoal.length, 7)
      .setValues(tabelSoal);

    // Format dashboard
    dashboard.getRange("A5:B5")
      .setFontWeight("bold")
      .setBackground("#d9ead3");

    dashboard.getRange("A37:G37")
      .setFontWeight("bold")
      .setBackground("#d9ead3");

    dashboard.getRange(21, 1, 1, 7)
      .setFontWeight("bold")
      .setBackground("#cfe2f3");

    dashboard.getRange("B9:B11").setNumberFormat("0.00");
    dashboard.getRange("B15").setNumberFormat("0.00");

    dashboard.getRange("D6:E7").setHorizontalAlignment("center");

    dashboard.autoResizeColumns(1, 8);
    dashboard.setFrozenRows(5);

    ui.alert(
      "DASHBOARD BERHASIL DIBUAT",
      "Dashboard untuk ujian " + ringkasan[1] +
      " sudah tersedia di sheet DASHBOARD_UJIAN.",
      ui.ButtonSet.OK
    );

  } catch (error) {
    ui.alert("PEMBUATAN DASHBOARD GAGAL:\n" + error.message);
  }
}




function mulaiImporUjianLama() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty("PROGRES_IMPOR_UJIAN")) {
    ui.alert("Masih ada proses impor. Jalankan lanjutkanImporUjianLama terlebih dahulu.");
    return;
  }

  const jawaban = ui.prompt(
    "Impor Ujian Lama",
    "Masukkan URL Google Form:",
    ui.ButtonSet.OK_CANCEL
  );

  if (jawaban.getSelectedButton() !== ui.Button.OK) return;

  const formId = jawaban.getResponseText().trim().match(/[-\w]{25,}/)?.[0];

  if (!formId) {
    ui.alert("URL Form tidak valid.");
    return;
  }

  const bank = ss.getSheetByName("BANK_SOAL");
  const daftar = ss.getSheetByName("DAFTAR_UJIAN");
  const pemetaan = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");

  if (!bank || !daftar || !pemetaan ||
      !ss.getSheetByName("Form Responses 7")) {
    ui.alert("Sheet yang diperlukan tidak ditemukan.");
    return;
  }

  try {
    const lastRow = daftar.getLastRow();

    if (lastRow > 1) {
      const ids = daftar.getRange(2, 3, lastRow - 1, 1)
        .getValues().flat().map(String);

      if (ids.includes(formId)) {
        ui.alert("Form ini sudah terdaftar.");
        return;
      }
    }

    const form = FormApp.openById(formId);
    const items = form.getItems();

    const soalForm = [];

    items.forEach((item, index) => {
      const judul = item.getTitle() || "";
      const cocok = judul.match(/8WG-\d{3}/);

      if (cocok) {
        soalForm.push({
          index: index,
          idSoal: cocok[0]
        });
      }
    });

    if (soalForm.length !== 18) {
      ui.alert("Ditemukan " + soalForm.length +
        " soal. Seharusnya 18. Proses dibatalkan.");
      return;
    }

    const lastBank = bank.getLastRow();
    const headers = bank.getRange(1, 1, 1, bank.getLastColumn())
      .getValues()[0];

    const kolom = {};
    headers.forEach((h, i) => {
      kolom[String(h).trim().toUpperCase()] = i + 1;
    });

    const wajib = [
      "ID_SOAL", "JENIS_SOAL", "MATERI",
      "LEVEL_BLOOM", "LEVEL_KOGNITIF",
      "KUNCI_JAWABAN", "JAWABAN_DITERIMA",
      "RUBRIK", "POIN"
    ];

    for (const nama of wajib) {
      if (!kolom[nama]) {
        ui.alert("Kolom " + nama + " tidak ditemukan.");
        return;
      }
    }

    const bankIds = bank.getRange(
      2, kolom.ID_SOAL, lastBank - 1, 1
    ).getValues().flat().map(String);

    const belumAda = soalForm
      .filter(s => !bankIds.includes(s.idSoal))
      .map(s => s.idSoal);

    if (belumAda.length) {
      ui.alert("ID tidak ditemukan di BANK_SOAL:\n" +
        belumAda.join(", "));
      return;
    }

    const idUjian = "UJIAN-" + Utilities.getUuid();

    props.setProperty("PROGRES_IMPOR_UJIAN", JSON.stringify({
      formId: formId,
      idUjian: idUjian,
      namaUjian: form.getTitle(),
      indeksSoal: soalForm.map(s => s.index),
      next: 0
    }));

    ui.alert(
      "Persiapan berhasil.\n" +
      "Ditemukan 18 soal.\n\n" +
      "Selanjutnya jalankan fungsi lanjutkanImporUjianLama."
    );

  } catch (e) {
    ui.alert("Gagal memulai impor: " + e.message);
  }
}


function lanjutkanImporUjianLama() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const data = props.getProperty("PROGRES_IMPOR_UJIAN");

  if (!data) {
    ui.alert("Belum ada proses impor. Jalankan mulaiImporUjianLama dahulu.");
    return;
  }

  try {
    const state = JSON.parse(data);

    const form = FormApp.openById(state.formId);
    const items = form.getItems();

    const bank = ss.getSheetByName("BANK_SOAL");
    const daftar = ss.getSheetByName("DAFTAR_UJIAN");
    const pemetaan = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");

    const headers = bank.getRange(
      1, 1, 1, bank.getLastColumn()
    ).getValues()[0];

    const kolom = {};
    headers.forEach((h, i) => {
      kolom[String(h).trim().toUpperCase()] = i + 1;
    });

    const lastBank = bank.getLastRow();

    const bankIds = bank.getRange(
      2, kolom.ID_SOAL, lastBank - 1, 1
    ).getValues().flat().map(String);

    const batch = state.indeksSoal.slice(state.next, state.next + 3);
    const rows = [];

    batch.forEach((index) => {
      const item = items[index];
      const judul = item.getTitle();
      const idSoal = judul.match(/8WG-\d{3}/)[0];

      const posisi = bankIds.indexOf(idSoal);

      if (posisi < 0) {
        throw new Error("ID " + idSoal + " tidak ditemukan.");
      }

      const nomorBaris = posisi + 2;

      const nilai = {};
      Object.keys(kolom).forEach(nama => {
        nilai[nama] = bank.getRange(
          nomorBaris, kolom[nama]
        ).getValue();
      });

      rows.push([
        state.idUjian,
        idSoal,
        String(item.getId()),
        state.next + rows.length + 1,
        "",
        nilai.JENIS_SOAL,
        judul,
        Number(nilai.POIN) || 0,
        nilai.KUNCI_JAWABAN,
        nilai.JAWABAN_DITERIMA,
        nilai.RUBRIK,
        nilai.MATERI,
        nilai.LEVEL_BLOOM,
        nilai.LEVEL_KOGNITIF,
        "PERLU_VALIDASI"
      ]);
    });

    pemetaan.getRange(
      pemetaan.getLastRow() + 1,
      1,
      rows.length,
      15
    ).setValues(rows);

    state.next += batch.length;

    if (state.next >= 18) {
      daftar.appendRow([
        state.idUjian,
        state.namaUjian,
        state.formId,
        form.getPublishedUrl(),
        ss.getId(),
        "Form Responses 7",
        new Date(),
        "",
        18,
        "PERLU_VALIDASI"
      ]);

      props.deleteProperty("PROGRES_IMPOR_UJIAN");

      ui.alert(
        "IMPOR SELESAI.\n" +
        "Jumlah soal: 18\n" +
        "ID Ujian: " + state.idUjian
      );

    } else {
      props.setProperty(
        "PROGRES_IMPOR_UJIAN",
        JSON.stringify(state)
      );

      ui.alert(
        "Berhasil memproses " + state.next +
        " dari 18 soal.\n\n" +
        "Jalankan kembali lanjutkanImporUjianLama."
      );
    }

  } catch (e) {
    ui.alert("Impor gagal: " + e.message);
  }
}


function batalImporUjianLama() {
  PropertiesService.getScriptProperties()
    .deleteProperty("PROGRES_IMPOR_UJIAN");

  SpreadsheetApp.getUi().alert(
    "Progres impor dibatalkan. Data yang sudah terlanjur ditulis tidak dihapus."
  );
}


function daftarkanUjianLama() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const idUjian = "UJIAN-47c7c057-929c-4f0b-ac93-1404b6c41e83";

  const sheetDaftar = ss.getSheetByName("DAFTAR_UJIAN");
  const sheetPemetaan = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");

  if (!sheetDaftar || !sheetPemetaan) {
    ui.alert("Sheet tidak ditemukan.");
    return;
  }

  const data = sheetPemetaan.getDataRange().getValues();

  const soal = data.filter((row, i) =>
    i > 0 && String(row[0]).trim() === idUjian
  );

  if (soal.length !== 18) {
    ui.alert("Jumlah pemetaan bukan 18. Ditemukan: " + soal.length);
    return;
  }

  const daftar = sheetDaftar.getDataRange().getValues();

  if (daftar.some((row, i) =>
    i > 0 && String(row[0]).trim() === idUjian
  )) {
    ui.alert("ID ujian sudah terdaftar.");
    return;
  }

  const formId = "GANTI_DENGAN_FORM_ID";

  const form = FormApp.openById(formId);

  sheetDaftar.appendRow([
    idUjian,
    form.getTitle(),
    formId,
    form.getPublishedUrl(),
    ss.getId(),
    "Form Responses 7",
    new Date(),
    "",
    18,
    "PERLU_VALIDASI"
  ]);

  ui.alert("Ujian berhasil didaftarkan ke DAFTAR_UJIAN.");
}



function sinkronkanBobotSoal() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const input = ui.prompt(
    "SINKRONISASI BOBOT SOAL",
    "Masukkan ID_UJIAN:",
    ui.ButtonSet.OK_CANCEL
  );

  if (input.getSelectedButton() !== ui.Button.OK) return;

  const idUjian = input.getResponseText().trim();

  const daftar = ss.getSheetByName("DAFTAR_UJIAN");
  const pemetaan = ss.getSheetByName("PEMETAAN_SOAL_UJIAN");

  const dataUjian = daftar.getDataRange().getValues();
  const ujian = dataUjian.find(
    r => String(r[0]).trim() === idUjian
  );

  if (!ujian) {
    ui.alert("ID_UJIAN tidak ditemukan.");
    return;
  }

  const form = FormApp.openById(ujian[2]);
  const itemForm = form.getItems();

  const bobotForm = {};

  itemForm.forEach(item => {
    const jenis = item.getType();
    let poin = null;

    switch (jenis) {
      case FormApp.ItemType.MULTIPLE_CHOICE:
        poin = item.asMultipleChoiceItem().getPoints();
        break;

      case FormApp.ItemType.CHECKBOX:
        poin = item.asCheckboxItem().getPoints();
        break;

      case FormApp.ItemType.TEXT:
        poin = item.asTextItem().getPoints();
        break;

      case FormApp.ItemType.PARAGRAPH_TEXT:
        poin = item.asParagraphTextItem().getPoints();
        break;

      case FormApp.ItemType.LIST:
        poin = item.asListItem().getPoints();
        break;
    }

    if (poin !== null) {
      bobotForm[String(item.getId())] = poin;
    }
  });

  const data = pemetaan.getDataRange().getValues();

  let jumlahSinkron = 0;
  let tidakDitemukan = 0;
  let totalPoin = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== idUjian) continue;

    const itemId = String(data[i][2]);
    const bobot = bobotForm[itemId];

    if (bobot === undefined) {
      tidakDitemukan++;
      continue;
    }

    pemetaan.getRange(i + 1, 8).setValue(bobot);

    jumlahSinkron++;
    totalPoin += bobot;
  }

  ui.alert(
    "Sinkronisasi selesai.\n\n" +
    "Soal berhasil disinkronkan: " + jumlahSinkron +
    "\nSoal tidak ditemukan: " + tidakDitemukan +
    "\nTotal poin maksimal: " + totalPoin
  );
}

function buatGoogleForm() {
  const html = HtmlService
    .createHtmlOutputFromFile("FormUjian")
    .setWidth(760)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    "BUAT GOOGLE FORM BARU"
  );
}