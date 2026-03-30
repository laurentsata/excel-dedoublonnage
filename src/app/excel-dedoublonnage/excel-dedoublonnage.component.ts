import { CommonModule } from '@angular/common';
import { Component, ElementRef, ViewChild } from '@angular/core';
import * as XLSX from 'xlsx';

interface ExcelRow {
  [key: string]: any;
}

interface DuplicateInfo {
  siret: string;
  nom: string;
  firstLine: number;
  duplicateLines: number[];
}

interface SiretErrorInfo {
  siret: string;
  nom: string;
  lineNumber: number;
  length: number;
  reason: string;
}

interface FieldErrorInfo {
  lineNumber: number;
  field: string;
  originalValue: string;
  reason: string;
  nom: string;
  siret: string;
}

@Component({
  selector: 'app-excel-dedoublonnage',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './excel-dedoublonnage.component.html',
  styleUrls: ['./excel-dedoublonnage.component.scss']
})
export class ExcelDedoublonnageComponent {
  originalRows: ExcelRow[] = [];
  cleanedRows: ExcelRow[] = [];
  duplicateInfos: DuplicateInfo[] = [];
  siretErrors: SiretErrorInfo[] = [];
  fieldErrors: FieldErrorInfo[] = [];
  @ViewChild('fileInput') fileInputRef!: ElementRef<HTMLInputElement>;

  fileName = '';
  sheetName = '';
  totalRows = 0;
  totalDuplicatesRemoved = 0;
  totalValidCleanedRows = 0;
  totalSiretErrors = 0;
  totalFieldErrors = 0;
  isFileLoaded = false;

  /**
   * Cette méthode est appelée quand l'utilisateur sélectionne un fichier Excel.
   * Elle lit le fichier puis transforme les données en objets exploitables.
   */
  onFileChange(event: Event): void {
    const input = event.target as HTMLInputElement;

    if (!input.files || input.files.length === 0) {
      return;
    }

    const file = input.files[0];
    this.fileName = file.name;

    const reader = new FileReader();

    reader.onload = (e: ProgressEvent<FileReader>) => {
      const data = e.target?.result;

      if (!data) {
        return;
      }

      // Cette ligne lit le classeur Excel importé
      const workbook = XLSX.read(data, { type: 'array' });

      // Cette ligne sélectionne la première feuille du fichier
      this.sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[this.sheetName];

      // Cette ligne lit la feuille en brut sous forme de tableau 2D
      const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: ''
      });

      // Cette ligne reconstruit les objets lignes depuis les données brutes
      this.originalRows = this.buildRowsFromRawData(rawData);

      // Cette ligne compte uniquement les lignes contenant un SIRET
      this.totalRows = this.originalRows.filter(row => {
        const siret = this.normalizeSiret(this.findSiretValue(row));
        return !!siret;
      }).length;

      // Cette ligne lance le traitement de suppression des doublons
      this.processRows();

      this.isFileLoaded = true;
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * Cette méthode reconstruit les lignes du fichier Excel.
   * Hypothèse actuelle :
   * - ligne 1 Excel = ligne parasite
   * - ligne 2 Excel = vraie ligne d'entête
   * - ligne 3+ Excel = données
   */
 buildRowsFromRawData(rawData: any[][]): ExcelRow[] {
  if (!rawData || rawData.length < 2) {
    return [];
  }

  // Cette ligne récupère la vraie ligne d'entête
  const headerRow = rawData[1];

  // Cette ligne récupère toutes les lignes de données
  const dataRows = rawData.slice(2);

  return dataRows.map((row: any[], dataIndex: number) => {
    const rowObject: ExcelRow = {};

    headerRow.forEach((headerCell: any, index: number) => {
      // Cette ligne sécurise le nom de colonne
      const columnName = String(headerCell ?? '').trim();

      // Cette condition ignore les colonnes sans nom
      if (!columnName) {
        return;
      }

      // Cette ligne reconstruit l'objet avec le bon nom de colonne
      rowObject[columnName] = row[index] ?? '';
    });

    // Cette ligne mémorise le vrai numéro de ligne Excel d'origine
    rowObject['__excelLineNumber'] = dataIndex + 3;

    return rowObject;
  });
}

  /**
   * Cette méthode traite les lignes importées :
   * - doublon = même SIRET + même NOM
   * - on garde uniquement la première occurrence
   */
  processRows(): void {
    const seenEntries = new Map<string, number>();
    const duplicatesMap = new Map<string, DuplicateInfo>();
    const cleaned: ExcelRow[] = [];

    this.originalRows.forEach((row, index) => {
      // Cette ligne récupère le SIRET
      const rawSiret = this.findSiretValue(row);

      // Cette ligne récupère le NOM
      const rawNom = this.findNomValue(row);

      // Cette ligne normalise le SIRET
      const normalizedSiret = this.normalizeSiret(rawSiret);

      // Cette ligne normalise le NOM
      const normalizedNom = this.normalizeText(rawNom);

      // Cette ligne calcule le vrai numéro de ligne Excel
      // Cette ligne récupère le vrai numéro de ligne Excel d'origine
      const excelLineNumber = Number(row['__excelLineNumber'] ?? index + 3);

      // Cette condition ignore les lignes sans SIRET
      if (!normalizedSiret) {
        return;
      }

      // Cette condition garde les lignes sans NOM
      if (!normalizedNom) {
        cleaned.push(row);
        return;
      }

      // Cette ligne construit la clé métier de dédoublonnage
      const duplicateKey = `${normalizedSiret}||${normalizedNom}`;

      // Cette condition conserve la première occurrence
      if (!seenEntries.has(duplicateKey)) {
        seenEntries.set(duplicateKey, excelLineNumber);
        cleaned.push(row);
        return;
      }

      // Cette ligne récupère la première ligne où la valeur a été vue
      const firstLine = seenEntries.get(duplicateKey)!;

      // Cette condition initialise le suivi du doublon si nécessaire
      if (!duplicatesMap.has(duplicateKey)) {
        duplicatesMap.set(duplicateKey, {
          siret: normalizedSiret,
          nom: normalizedNom,
          firstLine,
          duplicateLines: []
        });
      }

      // Cette ligne mémorise la ligne supprimée
      duplicatesMap.get(duplicateKey)?.duplicateLines.push(excelLineNumber);
    });

    this.cleanedRows = cleaned;
    this.duplicateInfos = Array.from(duplicatesMap.values());

    // Cette ligne compte uniquement les lignes conservées avec SIRET
    this.totalValidCleanedRows = this.cleanedRows.filter(row => {
      const siret = this.normalizeSiret(this.findSiretValue(row));
      return !!siret;
    }).length;

    // Cette ligne calcule le nombre réel de doublons supprimés
    this.totalDuplicatesRemoved = this.totalRows - this.totalValidCleanedRows;

    // Cette ligne analyse les erreurs de longueur sur les SIRET après dédoublonnage
    this.analyzeSiretErrors();

    // Cette ligne analyse les erreurs métier sur sexe et actif
    this.analyzeFieldErrors();
  }

  /**
   * Cette méthode analyse les SIRET après dédoublonnage
   * et repère ceux dont la longueur est différente de 14.
   */
  analyzeSiretErrors(): void {
    this.siretErrors = this.cleanedRows
      .map((row, index) => {
        // Cette ligne récupère le SIRET de la ligne courante
        const siret = this.normalizeSiret(this.findSiretValue(row));

        // Cette ligne récupère le NOM pour l'affichage
        const nom = String(this.findNomValue(row) ?? '').trim();

        // Cette ligne calcule le numéro de ligne dans le tableau nettoyé
        const lineNumber = Number(row['__excelLineNumber'] ?? index + 3);

        // Cette condition ignore les lignes sans SIRET
        if (!siret) {
          return null;
        }

        // Cette condition ignore les SIRET corrects de 14 chiffres
        if (siret.length === 14) {
          return null;
        }

        // Cette ligne construit le motif de l'erreur
        const reason =
          siret.length < 14
            ? 'Inférieur à 14 chiffres'
            : 'Supérieur à 14 chiffres';

        return {
          siret,
          nom,
          lineNumber,
          length: siret.length,
          reason
        } as SiretErrorInfo;
      })
      .filter((item): item is SiretErrorInfo => item !== null);

    // Cette ligne calcule le nombre total d'erreurs SIRET
    this.totalSiretErrors = this.siretErrors.length;
  }

  /**
   * Cette méthode analyse les champs métier après dédoublonnage
   * et repère les erreurs sur sexe et actif.
   */
  analyzeFieldErrors(): void {
    this.fieldErrors = this.cleanedRows.flatMap((row, index) => {
      const errors: FieldErrorInfo[] = [];

      // Cette ligne calcule le numéro de ligne dans le tableau nettoyé
      const lineNumber = Number(row['__excelLineNumber'] ?? index + 3);

      // Cette ligne récupère les infos utiles pour l'affichage
      const siret = this.normalizeSiret(this.findSiretValue(row));
      const nom = String(this.findNomValue(row) ?? '').trim();

      // Cette ligne lit la valeur brute du sexe
      const rawSexe = String(this.findSexeValue(row) ?? '').trim();

      // Cette ligne lit la valeur brute du champ actif
      const rawActif = String(this.findActifValue(row) ?? '').trim();

      // Cette ligne tente de normaliser le sexe
      const normalizedSexe = this.normalizeSexe(rawSexe);

      // Cette ligne tente de normaliser actif
      const normalizedActif = this.normalizeActif(rawActif);

      // Cette condition détecte une erreur sur le sexe
      if (rawSexe && !normalizedSexe) {
        errors.push({
          lineNumber,
          field: 'sexe',
          originalValue: rawSexe,
          reason: 'Valeur sexe invalide',
          nom,
          siret
        });
      }

      // Cette condition détecte une erreur sur actif
      if (rawActif && !normalizedActif) {
        errors.push({
          lineNumber,
          field: 'actif',
          originalValue: rawActif,
          reason: 'Valeur actif invalide',
          nom,
          siret
        });
      }

      return errors;
    });

    // Cette ligne calcule le nombre total d'erreurs métier
    this.totalFieldErrors = this.fieldErrors.length;
  }

  /**
   * Cette méthode recherche la valeur du SIRET.
   */
  findSiretValue(row: ExcelRow): unknown {
    const possibleKeys = [
      'SIRET',
      'Siret',
      'siret',
      'Numéro SIRET',
      'Numero SIRET'
    ];

    for (const key of possibleKeys) {
      if (key in row) {
        return row[key];
      }
    }

    return '';
  }

  /**
   * Cette méthode recherche la valeur du NOM.
   */
  findNomValue(row: ExcelRow): unknown {
    const possibleKeys = ['Nom', 'NOM', 'nom'];

    for (const key of possibleKeys) {
      if (key in row) {
        return row[key];
      }
    }

    return '';
  }

  /**
   * Cette méthode recherche la valeur de l'email.
   */
  findEmailValue(row: ExcelRow): string {
    const possibleKeys = ['Email', 'EMAIL', 'email', 'Mail', 'MAIL', 'mail'];

    for (const key of possibleKeys) {
      if (key in row) {
        return String(row[key] ?? '').trim();
      }
    }

    return '';
  }

  /**
   * Cette méthode recherche la valeur du prénom.
   */
  findPrenomValue(row: ExcelRow): string {
    const possibleKeys = ['Prénom', 'Prenom', 'PRENOM', 'prenom', 'prénom'];

    for (const key of possibleKeys) {
      if (key in row) {
        return String(row[key] ?? '').trim();
      }
    }

    return '';
  }

  /**
   * Cette méthode recherche la valeur du sexe.
   */
  findSexeValue(row: ExcelRow): string {
    const possibleKeys = ['Sexe', 'SEXE', 'sexe'];

    for (const key of possibleKeys) {
      if (key in row) {
        return String(row[key] ?? '').trim();
      }
    }

    return '';
  }

  /**
   * Cette méthode recherche la valeur du champ actif.
   */
  findActifValue(row: ExcelRow): string {
    const possibleKeys = ['Actif', 'ACTIF', 'actif'];

    for (const key of possibleKeys) {
      if (key in row) {
        return String(row[key] ?? '').trim();
      }
    }

    return '';
  }

  /**
   * Cette méthode recherche la valeur du rôle.
   */
  findRoleValue(row: ExcelRow): string {
    const possibleKeys = ['Role', 'Rôle', 'ROLE', 'RÔLE', 'role', 'rôle'];

    for (const key of possibleKeys) {
      if (key in row) {
        return String(row[key] ?? '').trim();
      }
    }

    return '';
  }

  /**
   * Cette méthode normalise un SIRET.
   */
  normalizeSiret(value: unknown): string {
    return String(value ?? '')
      .replace(/\s+/g, '')
      .trim();
  }

  /**
   * Cette méthode normalise un texte.
   */
  normalizeText(value: unknown): string {
    return String(value ?? '')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }
/**
 * Cette méthode normalise la valeur du sexe.
 */
normalizeSexe(value: unknown): string {
  const normalized = String(value ?? '')
    .trim()
    .toLowerCase();

  // Cette condition accepte une valeur déjà normalisée pour femme
  if (normalized === '2') {
    return '2';
  }

  // Cette condition accepte une valeur déjà normalisée pour homme
  if (normalized === '1') {
    return '1';
  }

  // Cette condition gère les valeurs féminines
  if (
    normalized === 'femme' ||
    normalized === 'féminin' ||
    normalized === 'feminin' ||
    normalized === 'F'
  ) {
    return '2';
  }

  // Cette condition gère les valeurs masculines
  if (
    normalized === 'homme' ||
    normalized === 'masculin' ||
    normalized === 'M'
  ) {
    return '1';
  }

  // Cette ligne retourne vide si la valeur est invalide
  return '';
}

/**
 * Cette méthode normalise la valeur actif.
 */
  normalizeActif(value: unknown): string {
  const normalized = String(value ?? '')
    .trim()
    .toLowerCase();

  // Cette condition accepte une valeur déjà normalisée active
  if (normalized === '1') {
    return '1';
  }

  // Cette condition accepte une valeur déjà normalisée inactive
  if (normalized === '0') {
    return '0';
  }

  // Cette condition transforme oui en 1
  if (normalized === 'oui') {
    return '1';
  }

  // Cette condition transforme non en 0
  if (normalized === 'non') {
    return '0';
  }

  // Cette ligne retourne vide si la valeur est invalide
  return '';
}

  /**
   * Cette méthode transforme une ligne source vers le format CSV final attendu.
   */
  mapRowForCsvExport(row: ExcelRow): ExcelRow {
    return {
      // Cette ligne mappe le SIRET vers la colonne finale "siret"
      siret: this.normalizeSiret(this.findSiretValue(row)),

      // Cette ligne mappe l'email vers la colonne finale "email"
      email: this.findEmailValue(row),

      // Cette ligne mappe le nom vers la colonne finale "nom"
      nom: this.findNomValue(row),

      // Cette ligne mappe le prénom vers la colonne finale "prenom"
      prenom: this.findPrenomValue(row),

      // Cette ligne normalise le sexe vers 1 ou 2
      sexe: this.normalizeSexe(this.findSexeValue(row)),

      // Cette ligne normalise actif vers 1 ou 0
      actif: this.normalizeActif(this.findActifValue(row)),

      // Cette ligne mappe le rôle vers la colonne finale "role"
      role: this.findRoleValue(row)
    };
  }

  /**
   * Cette méthode exporte les données nettoyées au format CSV UTF-8
   * avec la structure exacte attendue.
   */
  downloadCleanedFile(): void {
    if (!this.cleanedRows.length) {
      return;
    }

    // Cette ligne conserve uniquement les lignes avec SIRET
    const rowsToExport = this.cleanedRows.filter(row => {
      const siret = this.normalizeSiret(this.findSiretValue(row));
      return !!siret;
    });

    if (!rowsToExport.length) {
      return;
    }

    // Cette ligne transforme les lignes source vers le format final attendu
    const normalizedExportRows = rowsToExport.map(row => this.mapRowForCsvExport(row));

    // Cette ligne impose l'ordre exact des colonnes du CSV final
    const headers = ['siret', 'email', 'nom', 'prenom', 'sexe', 'actif', 'role'];

    // Cette ligne construit le contenu CSV final
    const csvContent = this.convertRowsToCsv(normalizedExportRows, headers);

    // Cette ligne ajoute le BOM UTF-8 pour Excel
    const bom = '\uFEFF';

    // Cette ligne crée le fichier CSV en mémoire
    const blob = new Blob([bom + csvContent], {
      type: 'text/csv;charset=utf-8;'
    });

    // Cette ligne prépare le téléchargement
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = this.buildOutputFileName();

    // Cette ligne déclenche le téléchargement
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Cette ligne nettoie l'URL temporaire
    window.URL.revokeObjectURL(url);
  }

  /**
   * Cette méthode convertit un tableau d'objets en contenu CSV.
   */
  convertRowsToCsv(rows: ExcelRow[], headers: string[]): string {
    const separator = ';';

    // Cette ligne construit l'entête du CSV dans l'ordre exact demandé
    const headerLine = headers.join(separator);

    // Cette ligne construit les lignes de données
    const dataLines = rows.map(row => {
      return headers
        .map(header => this.escapeCsvValue(row[header]))
        .join(separator);
    });

    // Cette ligne force Excel à utiliser le séparateur point-virgule
    return ['sep=;', headerLine, ...dataLines].join('\r\n');
  }

  /**
   * Cette méthode sécurise une valeur pour le CSV.
   */
  escapeCsvValue(value: unknown): string {
    const stringValue = String(value ?? '').trim();

    // Cette condition protège les valeurs sensibles pour le CSV
    if (
      stringValue.includes(';') ||
      stringValue.includes('"') ||
      stringValue.includes('\n') ||
      stringValue.includes('\r')
    ) {
      const escapedValue = stringValue.replace(/"/g, '""');
      return `"${escapedValue}"`;
    }

    return stringValue;
  }

  /**
 * Cette méthode indique s'il existe au moins une erreur bloquante.
 */
hasBlockingErrors(): boolean {
  return this.totalSiretErrors > 0 || this.totalFieldErrors > 0;
}

  /**
   * Cette méthode construit le nom du fichier CSV exporté.
   */
  buildOutputFileName(): string {
    const baseName = this.fileName.replace(/\.(xlsx|xls|csv)$/i, '');
    return `${baseName}_sans_doublons.csv`;
  }

/**
 * Cette méthode remet le composant à zéro.
 */
reset(): void {
  // Cette ligne vide les données d'origine
  this.originalRows = [];

  // Cette ligne vide les données nettoyées
  this.cleanedRows = [];

  // Cette ligne vide les doublons détectés
  this.duplicateInfos = [];

  // Cette ligne vide les erreurs SIRET
  this.siretErrors = [];

  // Cette ligne vide les erreurs métier
  this.fieldErrors = [];

  // Cette ligne réinitialise les infos fichier
  this.fileName = '';
  this.sheetName = '';

  // Cette ligne remet les compteurs à zéro
  this.totalRows = 0;
  this.totalValidCleanedRows = 0;
  this.totalDuplicatesRemoved = 0;
  this.totalSiretErrors = 0;
  this.totalFieldErrors = 0;

  // Cette ligne indique qu'aucun fichier n'est chargé
  this.isFileLoaded = false;

  // Cette ligne vide visuellement l'input file dans le DOM
  if (this.fileInputRef?.nativeElement) {
    this.fileInputRef.nativeElement.value = '';
  }
}


}