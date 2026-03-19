import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
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

  fileName = '';
  sheetName = '';
  totalRows = 0;
  totalDuplicatesRemoved = 0;
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

      // Lecture du fichier Excel
      const workbook = XLSX.read(data, { type: 'array' });

      // On prend la première feuille
      this.sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[this.sheetName];

      // Lecture brute de la feuille en tableau 2D
      const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: ''
      });

      // Reconstruction des lignes à partir de la vraie ligne d'entête
      this.originalRows = this.buildRowsFromRawData(rawData);
      this.totalRows = this.originalRows.length;

      // Traitement des doublons
      this.processRows();

      this.isFileLoaded = true;
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * Cette méthode reconstruit les lignes du fichier Excel.
   * Ici, on suppose :
   * - ligne 1 Excel = ligne parasite
   * - ligne 2 Excel = vraie ligne d'entête
   * - ligne 3+ Excel = données
   */
  buildRowsFromRawData(rawData: any[][]): ExcelRow[] {
    if (!rawData || rawData.length < 2) {
      return [];
    }

    // La vraie ligne d'entête est la ligne 2 Excel
    const headerRow = rawData[1];

    // Les données commencent à la ligne 3 Excel
    const dataRows = rawData.slice(2);

    return dataRows.map((row: any[]) => {
      const rowObject: ExcelRow = {};

      headerRow.forEach((headerCell: any, index: number) => {
        const columnName = String(headerCell ?? '').trim();

        if (!columnName) {
          return;
        }

        rowObject[columnName] = row[index] ?? '';
      });

      return rowObject;
    });
  }

  /**
   * Cette méthode traite les lignes importées :
   * - on considère qu'un doublon = même SIRET + même NOM
   * - on garde uniquement la première occurrence
   * - les autres sont supprimées
   */
  processRows(): void {
    const seenEntries = new Map<string, number>();
    const duplicatesMap = new Map<string, DuplicateInfo>();
    const cleaned: ExcelRow[] = [];

    this.originalRows.forEach((row, index) => {
      // Récupération du SIRET et du NOM
      const rawSiret = this.findSiretValue(row);
      const rawNom = this.findNomValue(row);

      // Normalisation des valeurs
      const normalizedSiret = this.normalizeSiret(rawSiret);
      const normalizedNom = this.normalizeText(rawNom);

      // Numéro de ligne Excel réel
      // index 0 = ligne 3 Excel
      const excelLineNumber = index + 3;

      // Si le SIRET ou le nom est vide, on garde la ligne
      // car on ne peut pas appliquer correctement la règle métier
      if (!normalizedSiret || !normalizedNom) {
        cleaned.push(row);
        return;
      }

      // Clé métier = SIRET + NOM
      const duplicateKey = `${normalizedSiret}||${normalizedNom}`;

      // Première occurrence => on garde
      if (!seenEntries.has(duplicateKey)) {
        seenEntries.set(duplicateKey, excelLineNumber);
        cleaned.push(row);
        return;
      }

      // Si on passe ici, c'est un doublon
      const firstLine = seenEntries.get(duplicateKey)!;

      if (!duplicatesMap.has(duplicateKey)) {
        duplicatesMap.set(duplicateKey, {
          siret: normalizedSiret,
          nom: normalizedNom,
          firstLine,
          duplicateLines: []
        });
      }

      // On mémorise la ligne supprimée
      duplicatesMap.get(duplicateKey)?.duplicateLines.push(excelLineNumber);

      // IMPORTANT :
      // On ne pousse pas la ligne dans cleaned,
      // donc elle est supprimée du fichier final
    });

    this.cleanedRows = cleaned;
    this.duplicateInfos = Array.from(duplicatesMap.values());
    this.totalDuplicatesRemoved = this.originalRows.length - this.cleanedRows.length;
  }

  /**
   * Cette méthode recherche la valeur du SIRET,
   * même si le nom exact de colonne varie.
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
   * Cette méthode recherche la valeur du NOM,
   * même si le nom exact de colonne varie.
   */
  findNomValue(row: ExcelRow): unknown {
    const possibleKeys = [
      'Nom',
      'NOM',
      'nom',
    ];

    for (const key of possibleKeys) {
      if (key in row) {
        return row[key];
      }
    }

    return '';
  }

  /**
   * Cette méthode normalise un SIRET :
   * - conversion en string
   * - suppression des espaces
   * - trim
   */
  normalizeSiret(value: unknown): string {
    return String(value ?? '')
      .replace(/\s+/g, '')
      .trim();
  }

  /**
   * Cette méthode normalise un texte :
   * - conversion en string
   * - trim
   * - suppression des espaces multiples
   * - passage en majuscules pour éviter les faux écarts
   */
  normalizeText(value: unknown): string {
    return String(value ?? '')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }

  /**
   * Cette méthode exporte le fichier nettoyé.
   */
  downloadCleanedFile(): void {
    if (!this.cleanedRows.length) {
      return;
    }

    const worksheet = XLSX.utils.json_to_sheet(this.cleanedRows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Donnees_nettoyees');
    XLSX.writeFile(workbook, this.buildOutputFileName());
  }

  /**
   * Cette méthode construit le nom du fichier de sortie.
   */
  buildOutputFileName(): string {
    const baseName = this.fileName.replace(/\.(xlsx|xls)$/i, '');
    return `${baseName}_sans_doublons.xlsx`;
  }

  /**
   * Cette méthode remet le composant à zéro.
   */
  reset(): void {
    this.originalRows = [];
    this.cleanedRows = [];
    this.duplicateInfos = [];
    this.fileName = '';
    this.sheetName = '';
    this.totalRows = 0;
    this.totalDuplicatesRemoved = 0;
    this.isFileLoaded = false;
  }
}