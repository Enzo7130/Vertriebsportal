/**
 * ============================================================
 * VERTRIEBSPORTAL / LERNPLATTFORM
 * ============================================================
 *
 * DATENBANK:
 * Spreadsheet "Vertragsannahmen"
 * Sheet "Senti Partner"
 *
 * KEIN Users-Sheet
 * KEIN Progress-Sheet
 *
 * Senti Partner Spalten:
 *
 * A  id
 * B  email
 * C  vorname
 * D  nachname
 * E  passwort_hash
 * F  salt
 * G  aktiv
 * H  reset_token
 * I  reset_gueltig_bis
 * J  erstellt_am
 * K  abgeschlossene_lessons
 *
 * ============================================================
 */


// ============================================================
// KONFIGURATION
// ============================================================

const CONFIG = {

  // Dein echtes Spreadsheet:
  SPREADSHEET_ID:
    '1KVrS2Q_cExWuUw5tezDdHd1SD_tSd43Rb2QIcoPNW2A',

  PARTNER_SHEET:
    'Senti Partner',

  CERTIFICATES_FOLDER:
    'Zertifikate',

  TOKEN_TTL_SECONDS:
    60 * 60 * 12,

  RESET_TTL_MINUTES:
    30,

  ITERATIONS:
    15000,

  SESSION_SECRET:
    'J7xQ9mK2vL8rT4nP6sW1yC5dF0hB3eR9uA7zX4kN'

};


// ============================================================
// KURSE
// ============================================================

const COURSE_CATALOG = [

  {
    id: 'cat-1',

    title:
      'Kategorie 1 – Grundlagen',

    lessons: [

      {
        id: '1-1',
        number: '1.1',
        title: 'Thema A: Unser Angebot',
        videoId: 'M7lc1UVf-VE',
        durationHint: 'Verpflichtendes Lernvideo'
      },

      {
        id: '1-2',
        number: '1.2',
        title: 'Thema B: Zielkunden',
        videoId: 'M7lc1UVf-VE',
        durationHint: 'Verpflichtendes Lernvideo'
      }

    ]

  },

  {
    id: 'cat-2',

    title:
      'Kategorie 2 – Vertriebspraxis',

    lessons: [

      {
        id: '2-1',
        number: '2.1',
        title: 'Thema A: Erstgespräch',
        videoId: 'M7lc1UVf-VE',
        durationHint: 'Verpflichtendes Lernvideo'
      },

      {
        id: '2-2',
        number: '2.2',
        title: 'Thema B: Bedarfsermittlung',
        videoId: 'M7lc1UVf-VE',
        durationHint: 'Verpflichtendes Lernvideo'
      }

    ]

  },

  {
    id: 'cat-3',

    title:
      'Kategorie 3 – Abschluss',

    lessons: [

      {
        id: '3-1',
        number: '3.1',
        title: 'Thema A: Angebot präsentieren',
        videoId: 'M7lc1UVf-VE',
        durationHint: 'Verpflichtendes Lernvideo'
      },

      {
        id: '3-2',
        number: '3.2',
        title: 'Thema B: Abschluss sichern',
        videoId: 'M7lc1UVf-VE',
        durationHint: 'Verpflichtendes Lernvideo'
      }

    ]

  }

];


// ============================================================
// WEB APP
// ============================================================

function doGet(e) {

  const template =
    HtmlService
      .createTemplateFromFile('Index');

  template.resetToken =
    (
      e &&
      e.parameter &&
      e.parameter.reset
    )
      ? e.parameter.reset
      : '';

  return template
    .evaluate()
    .setTitle('Vertriebsportal')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );

}


function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();

}


// ============================================================
// DATENBANK
// ============================================================

function getDatabase_() {

  return SpreadsheetApp
    .openById(
      CONFIG.SPREADSHEET_ID
    );

}


function getPartnerSheet_() {

  const ss =
    getDatabase_();

  const sheet =
    ss.getSheetByName(
      CONFIG.PARTNER_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Das Sheet "Senti Partner" wurde nicht gefunden.'
    );

  }

  return sheet;

}


// ============================================================
// SETUP
// ============================================================

function setup() {

  const sheet =
    getPartnerSheet_();

  const headers = [
    'id',
    'email',
    'vorname',
    'nachname',
    'passwort_hash',
    'salt',
    'aktiv',
    'reset_token',
    'reset_gueltig_bis',
    'erstellt_am',
    'abgeschlossene_lessons'
  ];

  /*
   * Nur Überschriften setzen,
   * wenn das Blatt leer ist.
   *
   * Es wird KEIN Users-Sheet
   * und KEIN Progress-Sheet erstellt.
   */

  if (sheet.getLastRow() === 0) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);

  }

  sheet.setFrozenRows(1);

  return {
    ok: true,
    spreadsheet:
      getDatabase_().getName(),
    sheet:
      CONFIG.PARTNER_SHEET
  };

}


// ============================================================
// REGISTRIERUNG
// ============================================================

function register(payload) {

  requireFields_(
    payload,
    [
      'firstName',
      'lastName',
      'email',
      'password'
    ]
  );

  const email =
    normalizeEmail_(
      payload.email
    );

  if (!isValidEmail_(email)) {

    throw new Error(
      'Bitte gib eine gültige E-Mail-Adresse ein.'
    );

  }

  if (
    String(payload.password).length < 10
  ) {

    throw new Error(
      'Das Passwort muss mindestens 10 Zeichen haben.'
    );

  }


  const lock =
    LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    const sheet =
      getPartnerSheet_();

    const existing =
      findUser_(email);


    /*
     * Wenn bereits ein Konto existiert,
     * darf kein zweites angelegt werden.
     */

    if (existing) {

      throw new Error(
        'Für diese E-Mail-Adresse besteht bereits ein Konto. Bitte melde dich an.'
      );

    }


    const salt =
      Utilities.getUuid() +
      Utilities.getUuid();


    const passwordHash =
      hashPassword_(
        payload.password,
        salt
      );


    /*
     * NEUER ACCOUNT:
     *
     * aktiv = nein
     *
     * Der Vertriebler kann sich erst
     * anmelden, wenn du im Sheet
     * manuell "ja" einträgst.
     */

    sheet.appendRow([

      Utilities.getUuid(),

      email,

      cleanText_(
        payload.firstName
      ),

      cleanText_(
        payload.lastName
      ),

      passwordHash,

      salt,

      'nein',

      '',

      '',

      new Date(),

      ''

    ]);


    /*
     * WICHTIG:
     *
     * Nach Registrierung KEIN Login.
     * Keine Session.
     * Kein Dashboard.
     */

    return {
      success: true
    };


  } finally {

    lock.releaseLock();

  }

}


// ============================================================
// LOGIN
// ============================================================

function login(payload) {

  requireFields_(
    payload,
    [
      'email',
      'password'
    ]
  );

  const email =
    normalizeEmail_(
      payload.email
    );


  const user =
    findUser_(email);


  if (!user) {

    throw new Error(
      'E-Mail-Adresse oder Passwort ist nicht korrekt.'
    );

  }


  /*
   * FREISCHALTUNG
   */

  if (
    user.aktiv !== 'ja'
  ) {

    throw new Error(
      'Dein Zugang wird noch geprüft. Sobald dein Konto freigeschaltet wurde, kannst du dich anmelden.'
    );

  }


  /*
   * PASSWORT
   */

  const enteredHash =
    hashPassword_(
      payload.password,
      user.salt
    );


  if (
    !safeEqual_(
      user.passwort_hash,
      enteredHash
    )
  ) {

    throw new Error(
      'E-Mail-Adresse oder Passwort ist nicht korrekt.'
    );

  }


  /*
   * SESSION
   */

  return {

    token:
      createSession_(
        user.email
      )

  };

}


// ============================================================
// REGISTRIERUNG PRÜFEN
// ============================================================

function checkRegistration(email) {

  const user =
    findUser_(
      normalizeEmail_(email)
    );

  return {

    registered:
      !!user

  };

}


// ============================================================
// DASHBOARD
// ============================================================

function getDashboard(token) {

  const user =
    requireSession_(token);


  const progress =
    getCompletedLessons_(
      user.email
    );


  return {

    user: {

      firstName:
        user.vorname,

      lastName:
        user.nachname,

      email:
        user.email

    },


    catalog:
      COURSE_CATALOG,


    completedLessonIds:
      progress

  };

}


// ============================================================
// LESSON ABSCHLIESSEN
// ============================================================

function markLessonComplete(
  token,
  lessonId
) {

  const user =
    requireSession_(token);


  const lesson =
    findLesson_(
      lessonId
    );


  if (!lesson) {

    throw new Error(
      'Dieses Training existiert nicht.'
    );

  }


  const sheet =
    getPartnerSheet_();


  const completed =
    getCompletedLessons_(
      user.email
    );


  /*
   * Bereits abgeschlossen?
   */

  if (
    completed.indexOf(
      lesson.id
    ) === -1
  ) {

    completed.push(
      lesson.id
    );


    saveCompletedLessons_(
      user.row,
      completed
    );

  }


  return {

    completedLessonIds:
      getCompletedLessons_(
        user.email
      )

  };

}


// ============================================================
// ZERTIFIKAT
// ============================================================

function createCertificate(
  token,
  categoryId
) {

  const user =
    requireSession_(token);


  const category =
    COURSE_CATALOG.find(
      function(category) {

        return (
          category.id ===
          categoryId
        );

      }
    );


  if (!category) {

    throw new Error(
      'Kategorie nicht gefunden.'
    );

  }


  const completed =
    getCompletedLessons_(
      user.email
    );


  const complete =
    category.lessons.every(
      function(lesson) {

        return (
          completed.indexOf(
            lesson.id
          ) !== -1
        );

      }
    );


  if (!complete) {

    throw new Error(
      'Bitte schließe zuerst alle Videos dieser Kategorie ab.'
    );

  }


  const date =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'dd.MM.yyyy'
    );


  const doc =
    DocumentApp.create(
      'Zertifikat – ' +
      user.vorname +
      ' ' +
      user.nachname +
      ' – ' +
      category.title
    );


  const body =
    doc.getBody();


  body.clear();


  body
    .appendParagraph(
      'ZERTIFIKAT'
    )
    .setHeading(
      DocumentApp
        .ParagraphHeading
        .TITLE
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  body
    .appendParagraph(
      '\nHiermit wird bestätigt, dass'
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  body
    .appendParagraph(
      user.vorname +
      ' ' +
      user.nachname
    )
    .setHeading(
      DocumentApp
        .ParagraphHeading
        .HEADING1
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  body
    .appendParagraph(
      '\ndie Kategorie'
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  body
    .appendParagraph(
      category.title
    )
    .setHeading(
      DocumentApp
        .ParagraphHeading
        .HEADING2
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  body
    .appendParagraph(
      '\nvollständig abgeschlossen hat.'
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  body
    .appendParagraph(
      '\nAusgestellt am ' +
      date
    )
    .setAlignment(
      DocumentApp
        .HorizontalAlignment
        .CENTER
    );


  doc.saveAndClose();


  const pdf =
    DriveApp
      .getFileById(
        doc.getId()
      )
      .getBlob()
      .getAs(
        MimeType.PDF
      )
      .setName(
        'Zertifikat-' +
        category.id +
        '-' +
        user.nachname +
        '.pdf'
      );


  const file =
    getCertificatesFolder_()
      .createFile(pdf);


  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );


  DriveApp
    .getFileById(
      doc.getId()
    )
    .setTrashed(true);


  return {

    url:
      'https://drive.google.com/uc?export=download&id=' +
      file.getId(),

    fileName:
      file.getName()

  };

}


// ============================================================
// PASSWORT VERGESSEN
// ============================================================

function requestPasswordReset(
  payload
) {

  if (
    !payload ||
    !payload.email
  ) {

    throw new Error(
      'Bitte gib deine E-Mail-Adresse ein.'
    );

  }


  const email =
    normalizeEmail_(
      payload.email
    );


  const user =
    findUser_(email);


  /*
   * Keine Information darüber
   * preisgeben, ob die E-Mail
   * existiert.
   */

  if (!user) {

    return {
      ok: true
    };

  }


  const token =
    Utilities.getUuid() +
    '-' +
    Utilities.getUuid();


  const expiry =
    new Date(
      Date.now() +
      CONFIG.RESET_TTL_MINUTES *
      60 *
      1000
    );


  const sheet =
    getPartnerSheet_();


  /*
   * Spalten H und I:
   *
   * H = reset_token
   * I = reset_gueltig_bis
   */

  sheet
    .getRange(
      user.row,
      8,
      1,
      2
    )
    .setValues([
      [
        hashToken_(token),
        expiry
      ]
    ]);


  const url =
    ScriptApp
      .getService()
      .getUrl() +
    '?reset=' +
    encodeURIComponent(token);


  MailApp.sendEmail({

    to:
      user.email,

    subject:
      'Passwort zurücksetzen – Vertriebsportal',

    htmlBody:
      '<p>Hallo ' +
      escapeHtml_(
        user.vorname
      ) +
      ',</p>' +

      '<p>du kannst dein Passwort über folgenden Link zurücksetzen:</p>' +

      '<p><a href="' +
      url +
      '">Passwort zurücksetzen</a></p>' +

      '<p>Der Link ist ' +
      CONFIG.RESET_TTL_MINUTES +
      ' Minuten gültig.</p>'

  });


  return {
    ok: true
  };

}


// ============================================================
// PASSWORT ZURÜCKSETZEN
// ============================================================

function resetPassword(
  payload
) {

  requireFields_(
    payload,
    [
      'token',
      'password'
    ]
  );


  if (
    String(
      payload.password
    ).length < 10
  ) {

    throw new Error(
      'Das Passwort muss mindestens 10 Zeichen haben.'
    );

  }


  const sheet =
    getPartnerSheet_();


  const data =
    sheet
      .getDataRange()
      .getValues();


  const tokenHash =
    hashToken_(
      payload.token
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const savedToken =
      String(
        data[i][7]
      );


    const expiry =
      data[i][8];


    if (

      safeEqual_(
        savedToken,
        tokenHash
      )

      &&

      expiry

      &&

      new Date(
        expiry
      ).getTime() >
      Date.now()

    ) {


      const salt =
        Utilities.getUuid() +
        Utilities.getUuid();


      const passwordHash =
        hashPassword_(
          payload.password,
          salt
        );


      /*
       * E = Passwort Hash
       * F = Salt
       * G = Aktiv bleibt unverändert
       * H = Reset Token löschen
       * I = Ablaufdatum löschen
       */

      sheet
        .getRange(
          i + 1,
          5,
          1,
          5
        )
        .setValues([
          [
            passwordHash,
            salt,
            data[i][6],
            '',
            ''
          ]
        ]);


      return {
        ok: true
      };

    }

  }


  throw new Error(
    'Dieser Link ist ungültig oder bereits abgelaufen.'
  );

}


// ============================================================
// USER SUCHEN
// ============================================================

function findUser_(email) {

  const sheet =
    getPartnerSheet_();


  const data =
    sheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const rowEmail =
      normalizeEmail_(
        data[i][1]
      );


    if (
      rowEmail === email
    ) {

      return {

        row:
          i + 1,

        id:
          data[i][0],

        email:
          data[i][1],

        vorname:
          data[i][2],

        nachname:
          data[i][3],

        passwort_hash:
          data[i][4],

        salt:
          data[i][5],

        aktiv:
          String(
            data[i][6]
          )
            .trim()
            .toLowerCase(),

        abgeschlossene_lessons:
          data[i][10]

      };

    }

  }


  return null;

}


// ============================================================
// LESSONS SPEICHERN
// ============================================================

function getCompletedLessons_(
  email
) {

  const user =
    findUser_(
      normalizeEmail_(email)
    );


  if (!user) {
    return [];
  }


  const value =
    user.abgeschlossene_lessons;


  if (!value) {
    return [];
  }


  /*
   * JSON-Format:
   *
   * ["1-1","1-2","2-1"]
   */

  try {

    const parsed =
      JSON.parse(
        String(value)
      );


    if (
      Array.isArray(parsed)
    ) {

      return parsed.map(
        String
      );

    }

  } catch (error) {

    /*
     * Falls eine alte Zeile
     * kein JSON enthält:
     *
     * vorsichtshalber leer starten.
     */

  }


  return [];

}


function saveCompletedLessons_(
  row,
  lessons
) {

  const sheet =
    getPartnerSheet_();


  const unique =
    Array.from(
      new Set(
        lessons.map(
          String
        )
      )
    );


  /*
   * Spalte K
   */

  sheet
    .getRange(
      row,
      11
    )
    .setValue(
      JSON.stringify(
        unique
      )
    );

}


// ============================================================
// LESSON SUCHEN
// ============================================================

function findLesson_(
  lessonId
) {

  for (
    const category
    of COURSE_CATALOG
  ) {

    for (
      const lesson
      of category.lessons
    ) {

      if (
        lesson.id ===
        lessonId
      ) {

        return Object.assign(
          {},
          lesson,
          {
            categoryId:
              category.id
          }
        );

      }

    }

  }


  return null;

}


// ============================================================
// SESSION
// ============================================================

function createSession_(
  email
) {

  const payload =
    Utilities.base64EncodeWebSafe(

      JSON.stringify({

        email:
          normalizeEmail_(
            email
          ),

        exp:
          Date.now() +
          CONFIG.TOKEN_TTL_SECONDS *
          1000

      })

    );


  return (

    payload +
    '.' +
    sign_(payload)

  );

}


function requireSession_(
  token
) {

  if (
    !token ||
    token.indexOf('.') === -1
  ) {

    throw new Error(
      'Deine Sitzung ist abgelaufen. Bitte melde dich erneut an.'
    );

  }


  const parts =
    token.split('.');


  if (
    parts.length !== 2
  ) {

    throw new Error(
      'Ungültige Sitzung.'
    );

  }


  if (
    !safeEqual_(
      sign_(parts[0]),
      parts[1]
    )
  ) {

    throw new Error(
      'Ungültige Sitzung.'
    );

  }


  let payload;


  try {

    payload =
      JSON.parse(

        Utilities
          .newBlob(
            Utilities.base64DecodeWebSafe(
              parts[0]
            )
          )
          .getDataAsString()

      );

  } catch (error) {

    throw new Error(
      'Ungültige Sitzung.'
    );

  }


  if (
    !payload.exp ||
    payload.exp < Date.now()
  ) {

    throw new Error(
      'Deine Sitzung ist abgelaufen. Bitte melde dich erneut an.'
    );

  }


  const user =
    findUser_(
      normalizeEmail_(
        payload.email
      )
    );


  if (!user) {

    throw new Error(
      'Dein Benutzerkonto wurde nicht gefunden.'
    );

  }


  if (
    user.aktiv !== 'ja'
  ) {

    throw new Error(
      'Dein Zugang ist nicht freigeschaltet.'
    );

  }


  return user;

}


// ============================================================
// SICHERHEIT / HASHING
// ============================================================

function sign_(
  value
) {

  return Utilities
    .base64EncodeWebSafe(

      Utilities
        .computeHmacSha256Signature(
          value,
          CONFIG.SESSION_SECRET
        )

    );

}


function hashPassword_(
  password,
  salt
) {

  let value =
    salt +
    ':' +
    String(password);


  for (
    let i = 0;
    i < CONFIG.ITERATIONS;
    i++
  ) {

    value =
      bytesToHex_(

        Utilities
          .computeDigest(
            Utilities.DigestAlgorithm.SHA_256,
            value
          )

      );

  }


  return value;

}


function hashToken_(
  value
) {

  return bytesToHex_(

    Utilities
      .computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        String(value)
      )

  );

}


function bytesToHex_(
  bytes
) {

  return bytes
    .map(
      function(b) {

        return (
          '0' +
          (
            b & 0xff
          ).toString(16)
        ).slice(-2);

      }
    )
    .join('');

}


function safeEqual_(
  a,
  b
) {

  return (

    String(a).length ===
    String(b).length

    &&

    String(a) ===
    String(b)

  );

}


// ============================================================
// HILFSFUNKTIONEN
// ============================================================

function normalizeEmail_(
  email
) {

  return String(
    email || ''
  )
    .trim()
    .toLowerCase();

}


function isValidEmail_(
  email
) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(email);

}


function cleanText_(
  value
) {

  return String(
    value || ''
  )
    .trim()
    .replace(
      /[<>]/g,
      ''
    )
    .slice(
      0,
      80
    );

}


function requireFields_(
  object,
  fields
) {

  if (
    !object ||
    fields.some(
      function(key) {
        return !object[key];
      }
    )
  ) {

    throw new Error(
      'Bitte fülle alle Pflichtfelder aus.'
    );

  }

}


function escapeHtml_(
  text
) {

  return String(
    text || ''
  ).replace(
    /[&<>'"]/g,
    function(c) {

      return {

        '&':
          '&amp;',

        '<':
          '&lt;',

        '>':
          '&gt;',

        "'":
          '&#39;',

        '"':
          '&quot;'

      }[c];

    }
  );

}


// ============================================================
// ZERTIFIKAT-ORDNER
// ============================================================

function getCertificatesFolder_() {

  const folders =
    DriveApp
      .getFoldersByName(
        CONFIG.CERTIFICATES_FOLDER
      );


  if (
    folders.hasNext()
  ) {

    return folders.next();

  }


  return DriveApp
    .createFolder(
      CONFIG.CERTIFICATES_FOLDER
    );

}
