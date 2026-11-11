
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

function buatGoogleForm() {

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

  // Meminta judul ujian
  const titleResponse = ui.prompt(
    "BUAT GOOGLE FORM",
    "Masukkan judul ujian:",
    ui.ButtonSet.OK_CANCEL
  );

  if (titleResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const formTitle = titleResponse.getResponseText().trim();

  if (!formTitle) {
    ui.alert("Judul ujian tidak boleh kosong.");
    return;
  }

  // Meminta deskripsi ujian
  const descResponse = ui.prompt(
    "DESKRIPSI UJIAN",
    "Masukkan petunjuk atau deskripsi ujian:",
    ui.ButtonSet.OK_CANCEL
  );

  if (descResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const formDescription = descResponse.getResponseText().trim();
  
  // Memilih pengacakan urutan soal
  const shuffleChoice = ui.alert(
    "ACAK URUTAN SOAL",
    "Apakah urutan pertanyaan akan diacak untuk setiap siswa?",
    ui.ButtonSet.YES_NO
  );

  const shuffleQuestions =
    shuffleChoice === ui.Button.YES;


  // Membaca kolom A-U
  const values = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 21
  ).getValues();

  // Mengambil semua soal yang tersedia
  const allQuestions = values.filter(row =>
    String(row[0]).trim() !== "" &&
    String(row[8]).trim() !== ""
  );

  // Memilih metode pengambilan soal
  const methodChoice = ui.alert(
    "METODE PEMILIHAN SOAL",
    "Pilih salah satu:\n\n" +
    "YA = Berdasarkan ID soal\n" +
    "TIDAK = Acak dari seluruh BANK_SOAL\n" +
    "BATAL = Berdasarkan kriteria tertentu",
    ui.ButtonSet.YES_NO_CANCEL
  );

  let questions = [];

  // METODE 1: BERDASARKAN ID
  if (methodChoice === ui.Button.YES) {

    const selectionResponse = ui.prompt(
      "PILIH SOAL",
      "Masukkan ID soal dipisahkan koma.\n\n" +
      "Contoh: SOAL001, SOAL002\n\n" +
      "Ketik * untuk semua soal.",
      ui.ButtonSet.OK_CANCEL
    );

    if (selectionResponse.getSelectedButton() !== ui.Button.OK) return;

    const selectionText = selectionResponse.getResponseText().trim();

    if (!selectionText) {
      ui.alert("ID soal belum diisi.");
      return;
    }

    const selectedIds = selectionText === "*"
      ? null
      : [...new Set(selectionText.split(",").map(id => id.trim()).filter(Boolean))];

    if (selectedIds !== null) {
      const existingIds = new Set(allQuestions.map(row => String(row[0]).trim()));
      const notFound = selectedIds.filter(id => !existingIds.has(id));

      if (notFound.length > 0) {
        ui.alert("ID SOAL TIDAK DITEMUKAN:\n\n" + notFound.join("\n"));
        return;
      }
    }

    questions = selectedIds === null
      ? [...allQuestions]
      : allQuestions.filter(row => selectedIds.includes(String(row[0]).trim()));

  // METODE 2: ACAK DARI SEMUA SOAL
  } else if (methodChoice === ui.Button.NO) {

    const numberResponse = ui.prompt(
      "JUMLAH SOAL ACAK",
      "Jumlah soal tersedia: " + allQuestions.length +
      "\nBerapa soal yang ingin diambil?",
      ui.ButtonSet.OK_CANCEL
    );

    if (numberResponse.getSelectedButton() !== ui.Button.OK) return;

    const numberOfQuestions = Number(numberResponse.getResponseText().trim());

    if (!Number.isInteger(numberOfQuestions) || numberOfQuestions < 1 || numberOfQuestions > allQuestions.length) {
      ui.alert("Jumlah soal tidak valid.");
      return;
    }

    questions = [...allQuestions];

    for (let i = questions.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [questions[i], questions[j]] = [questions[j], questions[i]];
    }

    questions = questions.slice(0, numberOfQuestions);

  // METODE 3: BERDASARKAN KRITERIA
  } else if (methodChoice === ui.Button.CANCEL) {

    const askCriterion = (title, message) => {
      const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
      if (response.getSelectedButton() !== ui.Button.OK) return null;
      return response.getResponseText().trim();
    };

    const subject = askCriterion("MATA PELAJARAN", "Masukkan mata pelajaran.\nKosongkan jika tidak digunakan.");
    if (subject === null) return;

    const grade = askCriterion("KELAS", "Masukkan kelas dengan angka Romawi, misalnya VIII.\nKosongkan jika tidak digunakan.");
    if (grade === null) return;

    const material = askCriterion("MATERI", "Masukkan kata kunci materi.\nKosongkan jika tidak digunakan.");
    if (material === null) return;

    const bloom = askCriterion("LEVEL BLOOM", "Masukkan C1 sampai C6.\nKosongkan jika tidak digunakan.");
    if (bloom === null) return;

    const cognitive = askCriterion("LEVEL KOGNITIF", "Masukkan L1, L2, atau L3.\nKosongkan jika tidak digunakan.");
    if (cognitive === null) return;

    const filteredQuestions = allQuestions.filter(row =>
      (!subject || String(row[2]).trim().toLowerCase() === subject.toLowerCase()) &&
      (!grade || String(row[3]).trim().toLowerCase() === grade.toLowerCase()) &&
      (!material || String(row[4]).toLowerCase().includes(material.toLowerCase())) &&
      (!bloom || String(row[6]).trim().toLowerCase() === bloom.toLowerCase()) &&
      (!cognitive || String(row[7]).trim().toLowerCase() === cognitive.toLowerCase())
    );

    if (filteredQuestions.length === 0) {
      ui.alert("Tidak ada soal yang sesuai dengan kriteria tersebut.");
      return;
    }

    const numberResponse = ui.prompt(
      "JUMLAH SOAL",
      "Soal yang sesuai kriteria: " + filteredQuestions.length +
      "\nBerapa soal yang ingin diambil secara acak?",
      ui.ButtonSet.OK_CANCEL
    );

    if (numberResponse.getSelectedButton() !== ui.Button.OK) return;

    const numberOfQuestions = Number(numberResponse.getResponseText().trim());

    if (!Number.isInteger(numberOfQuestions) || numberOfQuestions < 1 || numberOfQuestions > filteredQuestions.length) {
      ui.alert("Jumlah soal tidak valid. Masukkan angka 1 sampai " + filteredQuestions.length + ".");
      return;
    }

    questions = [...filteredQuestions];

    for (let i = questions.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [questions[i], questions[j]] = [questions[j], questions[i]];
    }

    questions = questions.slice(0, numberOfQuestions);

  } else {
    return;
  }

  if (questions.length === 0) {
    ui.alert("Tidak ditemukan soal yang dipilih.");
    return;
  }

  const plan = [];
  const errors = [];

  // Memvalidasi semua soal sebelum membuat Form
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
    
    Logger.log("ID SOAL: " + id + " | IMAGE ID: " + imageId);

    if (!Number.isInteger(points) || points < 0) {
      errors.push(
        id + ": POIN harus berupa bilangan bulat minimal 0."
      );
      return;
    }

    const options = [
      String(row[9]).trim(),
      String(row[10]).trim(),
      String(row[11]).trim(),
      String(row[12]).trim()
    ];

    const key = String(row[13]).trim()
      .split(",")
      .map(x => x.trim().toUpperCase())
      .filter(Boolean);

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
  id, type, question, points, material, objective, bloom, cognitive, material, objective, bloom, cognitive,
  description, options, key, imageId
});

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
        id, type, question, points, material, objective, bloom, cognitive, material, objective, bloom, cognitive,
        description, options, key
      });

    } else if (type === "MJ") {

      const rawPairs = String(row[15]).trim();

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
        id, type, question, points, material, objective, bloom, cognitive,
        description, pairs, rightOptions
      });

    } else if (type === "IS") {

      if (!String(row[14]).trim()) {
        errors.push(id + ": JAWABAN_DITERIMA kosong.");
        return;
      }

      plan.push({
        id, type, question, points, material, objective, bloom, cognitive, description
      });

    } else if (type === "UR") {

      if (!String(row[16]).trim()) {
        errors.push(id + ": RUBRIK kosong.");
        return;
      }

      plan.push({
        id, type, question, points, material, objective, bloom, cognitive, description
      });

    } else {

      errors.push(id + ": Jenis soal tidak dikenal.");

    }

  });

  if (errors.length > 0) {
    ui.alert(
      "GOOGLE FORM BELUM DIBUAT\n\n" +
      errors.join("\n")
    );
    return;
  }

  // Konfirmasi pembuatan Form
  const confirm = ui.alert(
    "KONFIRMASI",
    "Jumlah soal: " + plan.length +
    "\n\nGoogle Form baru akan dibuat." +
    "\nForm akan terhubung dengan spreadsheet ini." +
    "\n\nLanjutkan?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  let form = null;

  try {

    // Membuat Google Form
    form = FormApp.create(formTitle);

    form.setIsQuiz(true);
    form.setDescription(formDescription);
    form.setConfirmationMessage(
      "Jawaban kamu sudah berhasil dikirim."
    );
    form.setProgressBar(true);
    form.setShuffleQuestions(false);
    
    // Identitas peserta
    form.addTextItem()
      .setTitle("Nama Lengkap")
      .setRequired(true);

    form.addTextItem()
      .setTitle("Kelas")
      .setRequired(true);

    form.addTextItem()
      .setTitle("Nomor Absen")
      .setRequired(true);
      
    // Memisahkan halaman identitas dan halaman ujian
    form.addPageBreakItem()
      .setTitle("SOAL UJIAN")
      .setHelpText("Kerjakan semua pertanyaan dengan teliti.");

    // Membuat soal
    // Mengacak urutan soal melalui Apps Script
if (shuffleQuestions) {
  for (let i = plan.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [plan[i], plan[j]] = [plan[j], plan[i]];
  }
}

// Membuat soal
plan.forEach(q => {

      if (q.type === "PG") {
                // Menambahkan gambar jika tersedia
        if (q.imageId) {
          const imageFile = DriveApp.getFileById(q.imageId);
          const imageItem = form.addImageItem();

          imageItem
            .setTitle("Gambar soal " + q.id)
            .setImage(imageFile.getBlob())
            .setAlignment(FormApp.Alignment.CENTER)
            .setWidth(500);
        }
        const item = form.addMultipleChoiceItem();

        item.setTitle(q.id + " - " + q.question);

        // Deskripsi berada di dalam petak soal
        if (q.description) {
          item.setHelpText(q.description);
        }

        const choices = q.options.map((option, index) => {

  const letter = String.fromCharCode(65 + index);

  return item.createChoice(
    option,
    q.key[0] === letter
  );

});

        item.setChoices(choices);
        item.setPoints(q.points);
        item.setRequired(true);

      } else if (q.type === "PGK") {

        const item = form.addCheckboxItem();

        item.setTitle(q.id + " - " + q.question);

        if (q.description) {
          item.setHelpText(q.description);
        }

        const choices = q.options.map((option, index) => {

  const letter = String.fromCharCode(65 + index);

  return item.createChoice(
    option,
    q.key.includes(letter)
  );

});

        item.setChoices(choices);
        item.setPoints(q.points);
        item.setRequired(true);

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

          // Deskripsi berada di dalam setiap petak pasangan
          if (q.description) {
            item.setHelpText(q.description);
          }

          const choices = q.rightOptions.map(option =>
            item.createChoice(
              option,
              option === pair.right
            )
          );

          item.setChoices(choices);
          item.setPoints(
            quotient + (index < remainder ? 1 : 0)
          );
          item.setRequired(true);

        });

      } else if (q.type === "IS") {

        const item = form.addTextItem();

        item.setTitle(q.id + " - " + q.question);

        item.setHelpText(
          (q.description ? q.description + "\n\n" : "") +
          "Jawaban akan diperiksa oleh guru."
        );

        item.setPoints(q.points);
        item.setRequired(true);

      } else if (q.type === "UR") {

        const item = form.addParagraphTextItem();

        item.setTitle(q.id + " - " + q.question);

        item.setHelpText(
          (q.description ? q.description + "\n\n" : "") +
          "Jawaban akan diperiksa berdasarkan rubrik guru."
        );

        item.setPoints(q.points);
        item.setRequired(true);

      }

    });

    // Membuat kisi-kisi berdasarkan soal yang digunakan
buatKisiKisi_(plan);

    // Menghubungkan respons ke spreadsheet
    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      ss.getId()
    );

    // Menampilkan tautan
    ui.alert(
      "GOOGLE FORM BERHASIL DIBUAT\n\n" +
      "Jumlah soal: " + plan.length +
      "\n\nLINK UNTUK MENGEDIT:\n" +
      form.getEditUrl() +
      "\n\nLINK UNTUK PESERTA:\n" +
      form.getPublishedUrl()
    );

  } catch (error) {

    ui.alert(
      "GAGAL MEMBUAT GOOGLE FORM\n\n" +
      error.message +
      (form
        ? "\n\nPeriksa Drive. Form yang dibuat sebagian mungkin masih tersimpan."
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

