import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { BehaviorSubject, Observable } from 'rxjs';
import * as XLSX from 'xlsx';

export interface NameCard {
  id: string;
  name: string;
  contactNumber: string;
  zone: string;
  connection: string;
  location: string;
  info: string;
  infoDate: string;
  invite: string;
  inviteDate: string;
  plan: string;
  comments: string;
  notes: string;
}

export interface TrackerActivity {
  id: string;
  date: string; // YYYY-MM-DD
  day: string;
  salesMonth: string;
  weekNo: number | string;
  name: string;
  info: string;
  infoDate: string;
  invite: string;
  inviteDate: string;
  plan: string;
  newNumber: string;
  newConnection: string;
  newLocation: string;
  comments: string;
  notes: string;
}

export interface NewNumber {
  id: string;
  salesMonth: string;
  weekNo: number | string;
  dateAdded: string; // YYYY-MM-DD
  name: string;
  contactNumber: string;
  connection: string;
  location: string;
  comments: string;
  notes: string;
  addedToMain: string; // Yes / No
}

@Injectable({
  providedIn: 'root'
})
export class ExcelDataService {
  private nameCardsSubject = new BehaviorSubject<NameCard[]>([]);
  public nameCards$: Observable<NameCard[]> = this.nameCardsSubject.asObservable();

  private trackerActivitiesSubject = new BehaviorSubject<TrackerActivity[]>([]);
  public trackerActivities$: Observable<TrackerActivity[]> = this.trackerActivitiesSubject.asObservable();

  private newNumbersSubject = new BehaviorSubject<NewNumber[]>([]);
  public newNumbers$: Observable<NewNumber[]> = this.newNumbersSubject.asObservable();

  private isLoadedSubject = new BehaviorSubject<boolean>(false);
  public isLoaded$: Observable<boolean> = this.isLoadedSubject.asObservable();

  // Store other raw sheets to preserve them on export
  private rawReadme: any[][] = [];
  private rawDashboard: any[][] = [];
  private rawLists: any[][] = [];

  constructor(private http: HttpClient) {
    this.initializeData();
  }

  /**
   * Initializes data from LocalStorage or fetches the default Excel file.
   */
  private initializeData(): void {
    const cachedCards = localStorage.getItem('qnet_namecards');
    const cachedActivities = localStorage.getItem('qnet_activities');
    const cachedNewNumbers = localStorage.getItem('qnet_newnumbers');
    const cachedReadme = localStorage.getItem('qnet_raw_readme');
    const cachedDashboard = localStorage.getItem('qnet_raw_dashboard');
    const cachedLists = localStorage.getItem('qnet_raw_lists');

    if (cachedCards && cachedActivities && cachedNewNumbers) {
      this.nameCardsSubject.next(JSON.parse(cachedCards));
      this.trackerActivitiesSubject.next(JSON.parse(cachedActivities));
      this.newNumbersSubject.next(JSON.parse(cachedNewNumbers));
      
      if (cachedReadme) this.rawReadme = JSON.parse(cachedReadme);
      if (cachedDashboard) this.rawDashboard = JSON.parse(cachedDashboard);
      if (cachedLists) this.rawLists = JSON.parse(cachedLists);
      
      this.isLoadedSubject.next(true);
    } else {
      // Load default file
      this.http.get('/QNET_2026_Sales_Weekly_Tracker_Final.xlsx', { responseType: 'arraybuffer' })
        .subscribe({
          next: (buffer: ArrayBuffer) => {
            this.parseExcelBuffer(buffer);
          },
          error: (err) => {
            console.error('Error loading default Excel sheet. Initializing empty tables.', err);
            this.isLoadedSubject.next(true);
          }
        });
    }
  }

  /**
   * Parses Excel file buffer into structured data.
   */
  public parseExcelBuffer(buffer: ArrayBuffer): void {
    try {
      const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: false });

      // 1. Read Me sheet (preserve)
      if (workbook.Sheets['Read Me']) {
        this.rawReadme = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets['Read Me'], { header: 1 });
      }

      // 2. Dashboard sheet (preserve)
      if (workbook.Sheets['Dashboard']) {
        this.rawDashboard = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets['Dashboard'], { header: 1 });
      }

      // 3. Lists sheet (preserve)
      if (workbook.Sheets['Lists']) {
        this.rawLists = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets['Lists'], { header: 1 });
      }

      // 4. Main (NameCards) Sheet
      const nameCards: NameCard[] = [];
      if (workbook.Sheets['Main']) {
        const rows = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets['Main'], { header: 1 });
        const headers = rows[0] || [];
        
        // Find header indices
        const nameIdx = headers.indexOf('Name (DisplayName)');
        const numIdx = headers.indexOf('Contact Number');
        const zoneIdx = headers.indexOf('Zone');
        const connIdx = headers.indexOf('Connection');
        const locIdx = headers.indexOf('Location');
        const infoIdx = headers.indexOf('Info');
        const infoDateIdx = headers.indexOf('Info_Date');
        const inviteIdx = headers.indexOf('Invite');
        const inviteDateIdx = headers.indexOf('Invite_Date');
        const planIdx = headers.indexOf('Plan');
        const commIdx = headers.indexOf('Comments');
        const notesIdx = headers.indexOf('Notes');

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r || r.length === 0) continue;
          
          const name = r[nameIdx] ? String(r[nameIdx]).trim() : '';
          const contact = r[numIdx] ? String(r[numIdx]).trim() : '';
          
          // Skip placeholder/empty rows
          if (!name && !contact) continue;

          nameCards.push({
            id: 'nc_' + Math.random().toString(36).substr(2, 9),
            name,
            contactNumber: contact,
            zone: r[zoneIdx] ? String(r[zoneIdx]).trim() : '',
            connection: r[connIdx] ? String(r[connIdx]).trim() : '',
            location: r[locIdx] ? String(r[locIdx]).trim() : '',
            info: r[infoIdx] ? String(r[infoIdx]).trim() : '',
            infoDate: this.formatExcelDate(r[infoDateIdx]),
            invite: r[inviteIdx] ? String(r[inviteIdx]).trim() : '',
            inviteDate: this.formatExcelDate(r[inviteDateIdx]),
            plan: r[planIdx] ? String(r[planIdx]).trim() : '',
            comments: r[commIdx] ? String(r[commIdx]).trim() : '',
            notes: r[notesIdx] ? String(r[notesIdx]).trim() : ''
          });
        }
      }

      // 5. Weekly Tracker Sheet
      const activities: TrackerActivity[] = [];
      if (workbook.Sheets['Weekly Tracker']) {
        const rows = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets['Weekly Tracker'], { header: 1 });
        const headers = rows[0] || [];

        const dateIdx = headers.indexOf('Date');
        const dayIdx = headers.indexOf('Day');
        const monthIdx = headers.indexOf('Sales Month');
        const weekIdx = headers.indexOf('Week No');
        const nameIdx = headers.indexOf('Name');
        const infoIdx = headers.indexOf('Info');
        const infoDateIdx = headers.indexOf('Info_Date');
        const inviteIdx = headers.indexOf('Invite');
        const inviteDateIdx = headers.indexOf('Invite_Date');
        const planIdx = headers.indexOf('Plan');
        const newNumIdx = headers.indexOf('New Number');
        const newConnIdx = headers.indexOf('New Connection');
        const newLocIdx = headers.indexOf('New Location');
        const commIdx = headers.indexOf('Comments');
        const notesIdx = headers.indexOf('Notes');

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r || r.length === 0) continue;

          const name = r[nameIdx] ? String(r[nameIdx]).trim() : '';
          const info = r[infoIdx] ? String(r[infoIdx]).trim() : '';
          const invite = r[inviteIdx] ? String(r[inviteIdx]).trim() : '';
          const plan = r[planIdx] ? String(r[planIdx]).trim() : '';

          // Only keep activities that are not placeholders (either has a name or some action log)
          if (!name && !info && !invite && !plan && !r[newNumIdx]) continue;

          activities.push({
            id: 'ta_' + Math.random().toString(36).substr(2, 9),
            date: this.formatExcelDate(r[dateIdx]),
            day: r[dayIdx] ? String(r[dayIdx]).trim() : '',
            salesMonth: r[monthIdx] ? String(r[monthIdx]).trim() : '',
            weekNo: r[weekIdx] !== undefined ? r[weekIdx] : '',
            name,
            info,
            infoDate: this.formatExcelDate(r[infoDateIdx]),
            invite,
            inviteDate: this.formatExcelDate(r[inviteDateIdx]),
            plan,
            newNumber: r[newNumIdx] ? String(r[newNumIdx]).trim() : '',
            newConnection: r[newConnIdx] ? String(r[newConnIdx]).trim() : '',
            newLocation: r[newLocIdx] ? String(r[newLocIdx]).trim() : '',
            comments: r[commIdx] ? String(r[commIdx]).trim() : '',
            notes: r[notesIdx] ? String(r[notesIdx]).trim() : ''
          });
        }
      }

      // 6. New Numbers Sheet
      const newNumbers: NewNumber[] = [];
      if (workbook.Sheets['New Numbers']) {
        const rows = XLSX.utils.sheet_to_json<any[]>(workbook.Sheets['New Numbers'], { header: 1 });
        const headers = rows[0] || [];

        const monthIdx = headers.indexOf('Sales Month');
        const weekIdx = headers.indexOf('Week No');
        const dateAddedIdx = headers.indexOf('Date Added');
        const nameIdx = headers.indexOf('Name');
        const numIdx = headers.indexOf('Contact Number');
        const connIdx = headers.indexOf('Connection');
        const locIdx = headers.indexOf('Location');
        const commIdx = headers.indexOf('Comments');
        const notesIdx = headers.indexOf('Notes');
        const addedIdx = headers.indexOf('Added to Main?');

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r || r.length === 0) continue;

          const name = r[nameIdx] ? String(r[nameIdx]).trim() : '';
          const contact = r[numIdx] ? String(r[numIdx]).trim() : '';

          if (!name && !contact) continue;

          newNumbers.push({
            id: 'nn_' + Math.random().toString(36).substr(2, 9),
            salesMonth: r[monthIdx] ? String(r[monthIdx]).trim() : '',
            weekNo: r[weekIdx] !== undefined ? r[weekIdx] : '',
            dateAdded: this.formatExcelDate(r[dateAddedIdx]),
            name,
            contactNumber: contact,
            connection: r[connIdx] ? String(r[connIdx]).trim() : '',
            location: r[locIdx] ? String(r[locIdx]).trim() : '',
            comments: r[commIdx] ? String(r[commIdx]).trim() : '',
            notes: r[notesIdx] ? String(r[notesIdx]).trim() : '',
            addedToMain: r[addedIdx] ? String(r[addedIdx]).trim() : 'No'
          });
        }
      }

      this.nameCardsSubject.next(nameCards);
      this.trackerActivitiesSubject.next(activities);
      this.newNumbersSubject.next(newNumbers);

      this.saveToLocalStorage();
      this.isLoadedSubject.next(true);

    } catch (e) {
      console.error('Failed to parse Excel array buffer:', e);
      this.isLoadedSubject.next(true);
    }
  }

  /**
   * Helper to format Excel dates (serial numbers or text strings) to YYYY-MM-DD.
   */
  public formatExcelDate(val: any): string {
    if (!val) return '';
    if (typeof val === 'number') {
      // Excel serial date to JS date
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    if (typeof val === 'string') {
      const trimmed = val.trim();
      // Check if it's already in YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
        return trimmed;
      }
      
      // Parse DD-MMM-YYYY (e.g. 01-Aug-2026)
      const parts = trimmed.split('-');
      if (parts.length === 3) {
        const day = parts[0].padStart(2, '0');
        const year = parts[2];
        const monthStr = parts[1].toLowerCase();
        const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        const monthIndex = months.findIndex(m => monthStr.startsWith(m));
        if (monthIndex !== -1 && year.length === 4) {
          const month = String(monthIndex + 1).padStart(2, '0');
          return `${year}-${month}-${day}`;
        }
      }
      
      // Try native JS parsing
      const parsed = Date.parse(trimmed);
      if (!isNaN(parsed)) {
        const d = new Date(parsed);
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      }
      return trimmed;
    }
    return String(val);
  }

  /**
   * Converts YYYY-MM-DD back into Excel serial number or readable text for sheet saving.
   * We will export standard DD-MMM-YYYY strings which Excel parses beautifully.
   */
  public formatToExcelDateString(dateStr: string): string {
    if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;
    const parts = dateStr.split('-');
    const year = parts[0];
    const monthIndex = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return `${String(day).padStart(2, '0')}-${months[monthIndex]}-${year}`;
  }

  private saveToLocalStorage(): void {
    localStorage.setItem('qnet_namecards', JSON.stringify(this.nameCardsSubject.value));
    localStorage.setItem('qnet_activities', JSON.stringify(this.trackerActivitiesSubject.value));
    localStorage.setItem('qnet_newnumbers', JSON.stringify(this.newNumbersSubject.value));
    localStorage.setItem('qnet_raw_readme', JSON.stringify(this.rawReadme));
    localStorage.setItem('qnet_raw_dashboard', JSON.stringify(this.rawDashboard));
    localStorage.setItem('qnet_raw_lists', JSON.stringify(this.rawLists));
  }

  // ==========================================
  // CRUD OPERATIONS: NAMECARDS
  // ==========================================

  public addNameCard(card: Omit<NameCard, 'id'>): void {
    const current = this.nameCardsSubject.value;
    const newCard: NameCard = {
      ...card,
      id: 'nc_' + Math.random().toString(36).substr(2, 9)
    };
    const updated = [newCard, ...current];
    this.nameCardsSubject.next(updated);
    this.saveToLocalStorage();
  }

  public editNameCard(card: NameCard): void {
    const current = this.nameCardsSubject.value;
    const updated = current.map(item => item.id === card.id ? card : item);
    this.nameCardsSubject.next(updated);
    this.saveToLocalStorage();
  }

  public deleteNameCard(id: string): void {
    const current = this.nameCardsSubject.value;
    const updated = current.filter(item => item.id !== id);
    this.nameCardsSubject.next(updated);
    this.saveToLocalStorage();
  }

  // ==========================================
  // CRUD OPERATIONS: TRACKER ACTIVITIES
  // ==========================================

  private syncActivityToNameCard(activity: TrackerActivity, previousName?: string): void {
    const targetName = (activity.name || previousName || '').trim();
    if (!targetName) return;

    const currentCards = this.nameCardsSubject.value;
    const normalizedTarget = targetName.toLowerCase();
    const nameCardIndex = currentCards.findIndex(card => (card.name || '').trim().toLowerCase() === normalizedTarget);
    const previousIndex = previousName
      ? currentCards.findIndex(card => (card.name || '').trim().toLowerCase() === previousName.trim().toLowerCase())
      : -1;

    const indexToUpdate = nameCardIndex !== -1 ? nameCardIndex : previousIndex;
    if (indexToUpdate === -1) return;

    const existingCard = currentCards[indexToUpdate];
    const updatedCards = [...currentCards];
    updatedCards[indexToUpdate] = {
      ...existingCard,
      name: activity.name?.trim() || existingCard.name,
      contactNumber: activity.newNumber || existingCard.contactNumber,
      location: activity.newLocation || existingCard.location,
      info: activity.info ?? existingCard.info,
      infoDate: activity.infoDate ?? existingCard.infoDate,
      invite: activity.invite ?? existingCard.invite,
      inviteDate: activity.inviteDate ?? existingCard.inviteDate,
      plan: activity.plan ?? existingCard.plan,
      comments: activity.comments ?? existingCard.comments,
      notes: activity.notes ?? existingCard.notes
    };

    this.nameCardsSubject.next(updatedCards);
    this.saveToLocalStorage();
  }

  public addActivity(act: Omit<TrackerActivity, 'id'>): void {
    const current = this.trackerActivitiesSubject.value;
    const newAct: TrackerActivity = {
      ...act,
      id: 'ta_' + Math.random().toString(36).substr(2, 9)
    };
    const updated = [newAct, ...current];
    this.trackerActivitiesSubject.next(updated);
    this.syncActivityToNameCard(newAct);
    this.saveToLocalStorage();
  }

  public editActivity(act: TrackerActivity): void {
    const current = this.trackerActivitiesSubject.value;
    const existingAct = current.find(item => item.id === act.id);
    const previousName = existingAct?.name || '';
    const updated = current.map(item => item.id === act.id ? act : item);
    this.trackerActivitiesSubject.next(updated);
    this.syncActivityToNameCard(act, previousName);
    this.saveToLocalStorage();
  }

  public deleteActivity(id: string): void {
    const current = this.trackerActivitiesSubject.value;
    const updated = current.filter(item => item.id !== id);
    this.trackerActivitiesSubject.next(updated);
    this.saveToLocalStorage();
  }

  // ==========================================
  // CRUD OPERATIONS: NEW NUMBERS
  // ==========================================

  public addNewNumber(num: Omit<NewNumber, 'id'>): void {
    const current = this.newNumbersSubject.value;
    const newNum: NewNumber = {
      ...num,
      id: 'nn_' + Math.random().toString(36).substr(2, 9)
    };

    // Add to New Numbers list
    const updated = [newNum, ...current];
    this.newNumbersSubject.next(updated);

    // Automatically promote to main NameCards if requested (or default to Yes as per preference)
    const shouldPromote = (newNum.addedToMain || 'Yes').toLowerCase() === 'yes';

    if (shouldPromote) {
      // Try to find existing card by contactNumber or name (case-insensitive)
      const existing = this.nameCardsSubject.value;
      const matchIndex = existing.findIndex(c => {
        const sameNumber = newNum.contactNumber && c.contactNumber && c.contactNumber.trim().toLowerCase() === newNum.contactNumber.trim().toLowerCase();
        const sameName = newNum.name && c.name && c.name.trim().toLowerCase() === newNum.name.trim().toLowerCase();
        return sameNumber || sameName;
      });

      const cardDetails: Omit<NameCard, 'id'> = {
        name: newNum.name || '',
        contactNumber: newNum.contactNumber || '',
        zone: newNum.connection || '',
        connection: newNum.connection || '',
        location: newNum.location || '',
        info: '',
        infoDate: '',
        invite: '',
        inviteDate: '',
        plan: 'No',
        comments: newNum.comments || '',
        notes: newNum.notes || ''
      };

      if (matchIndex !== -1) {
        // Update existing card with any new details
        const existingCard = existing[matchIndex];
        const updatedCard: NameCard = {
          ...existingCard,
          ...cardDetails
        };
        this.editNameCard(updatedCard);
      } else {
        this.addNameCard(cardDetails);
      }

      // Mark the NewNumber as added
      const updatedNums = this.newNumbersSubject.value.map(item => item.id === newNum.id ? { ...item, addedToMain: 'Yes' } : item);
      this.newNumbersSubject.next(updatedNums);
    }

    this.saveToLocalStorage();
  }

  public editNewNumber(num: NewNumber): void {
    const current = this.newNumbersSubject.value;
    const updated = current.map(item => item.id === num.id ? num : item);
    this.newNumbersSubject.next(updated);

    // If this NewNumber was already promoted (or should be), sync changes to NameCards
    const added = (num.addedToMain || 'No').toLowerCase() === 'yes';
    if (added) {
      // Find corresponding NameCard by contact number or name and update
      const cards = this.nameCardsSubject.value;
      const matchIndex = cards.findIndex(c => {
        const sameNumber = num.contactNumber && c.contactNumber && c.contactNumber.trim().toLowerCase() === num.contactNumber.trim().toLowerCase();
        const sameName = num.name && c.name && c.name.trim().toLowerCase() === num.name.trim().toLowerCase();
        return sameNumber || sameName;
      });

      if (matchIndex !== -1) {
        const existingCard = cards[matchIndex];
        const updatedCard: NameCard = {
          ...existingCard,
          name: num.name || existingCard.name,
          contactNumber: num.contactNumber || existingCard.contactNumber,
          connection: num.connection || existingCard.connection,
          location: num.location || existingCard.location,
          comments: num.comments || existingCard.comments,
          notes: num.notes || existingCard.notes
        };
        this.editNameCard(updatedCard);
      } else {
        // If not found, add as a new NameCard
        const cardDetails: Omit<NameCard, 'id'> = {
          name: num.name || '',
          contactNumber: num.contactNumber || '',
          zone: num.connection || '',
          connection: num.connection || '',
          location: num.location || '',
          info: '',
          infoDate: '',
          invite: '',
          inviteDate: '',
          plan: 'No',
          comments: num.comments || '',
          notes: num.notes || ''
        };
        this.addNameCard(cardDetails);
      }
    }

    this.saveToLocalStorage();
  }

  public deleteNewNumber(id: string): void {
    const current = this.newNumbersSubject.value;
    const updated = current.filter(item => item.id !== id);
    this.newNumbersSubject.next(updated);
    this.saveToLocalStorage();
  }

  /**
   * Promotes a NewNumber to the Main NameCards list and marks it as added.
   */
  public promoteToMain(numId: string, cardDetails: Omit<NameCard, 'id'>): void {
    // Add to Main NameCards
    this.addNameCard(cardDetails);

    // Update NewNumber status
    const currentNums = this.newNumbersSubject.value;
    const updatedNums = currentNums.map(item => {
      if (item.id === numId) {
        return { ...item, addedToMain: 'Yes' };
      }
      return item;
    });
    this.newNumbersSubject.next(updatedNums);
    this.saveToLocalStorage();
  }

  // ==========================================
  // EXPORT ENGINE
  // ==========================================

  /**
   * Rebuilds the workbook and downloads the excel sheet
   */
  public exportToExcel(): void {
    const wb = XLSX.utils.book_new();

    // 1. Read Me Sheet
    if (this.rawReadme.length > 0) {
      const ws = XLSX.utils.aoa_to_sheet(this.rawReadme);
      XLSX.utils.book_append_sheet(wb, ws, 'Read Me');
    } else {
      // Dummy Read Me
      const ws = XLSX.utils.aoa_to_sheet([['QNET SALES TRACKER EXPORT']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Read Me');
    }

    // 2. Dashboard Sheet
    if (this.rawDashboard.length > 0) {
      // We can output the dashboard raw rows as is
      const ws = XLSX.utils.aoa_to_sheet(this.rawDashboard);
      XLSX.utils.book_append_sheet(wb, ws, 'Dashboard');
    }

    // 3. Main Sheet
    const mainHeaders = [
      'Name (DisplayName)',
      'Contact Number',
      'Zone',
      'Connection',
      'Location',
      'Info',
      'Info_Date',
      'Invite',
      'Invite_Date',
      'Plan',
      'Comments',
      'Notes'
    ];
    const mainData = this.nameCardsSubject.value.map(card => [
      card.name,
      card.contactNumber,
      card.zone,
      card.connection,
      card.location,
      card.info,
      this.formatToExcelDateString(card.infoDate),
      card.invite,
      this.formatToExcelDateString(card.inviteDate),
      card.plan,
      card.comments,
      card.notes
    ]);
    const wsMain = XLSX.utils.aoa_to_sheet([mainHeaders, ...mainData]);
    XLSX.utils.book_append_sheet(wb, wsMain, 'Main');

    // 4. Weekly Tracker Sheet
    const trackerHeaders = [
      'Date',
      'Day',
      'Sales Month',
      'Week No',
      'Name',
      'Info',
      'Info_Date',
      'Invite',
      'Invite_Date',
      'Plan',
      'New Number',
      'New Connection',
      'New Location',
      'Comments',
      'Notes'
    ];
    const trackerData = this.trackerActivitiesSubject.value.map(act => [
      this.formatToExcelDateString(act.date),
      act.day,
      act.salesMonth,
      act.weekNo,
      act.name,
      act.info,
      this.formatToExcelDateString(act.infoDate),
      act.invite,
      this.formatToExcelDateString(act.inviteDate),
      act.plan,
      act.newNumber,
      act.newConnection,
      act.newLocation,
      act.comments,
      act.notes
    ]);
    const wsTracker = XLSX.utils.aoa_to_sheet([trackerHeaders, ...trackerData]);
    XLSX.utils.book_append_sheet(wb, wsTracker, 'Weekly Tracker');

    // 5. New Numbers Sheet
    const newNumHeaders = [
      'Sales Month',
      'Week No',
      'Date Added',
      'Name',
      'Contact Number',
      'Connection',
      'Location',
      'Comments',
      'Notes',
      'Added to Main?'
    ];
    const newNumData = this.newNumbersSubject.value.map(num => [
      num.salesMonth,
      num.weekNo,
      this.formatToExcelDateString(num.dateAdded),
      num.name,
      num.contactNumber,
      num.connection,
      num.location,
      num.comments,
      num.notes,
      num.addedToMain
    ]);
    const wsNewNum = XLSX.utils.aoa_to_sheet([newNumHeaders, ...newNumData]);
    XLSX.utils.book_append_sheet(wb, wsNewNum, 'New Numbers');

    // 6. Lists Sheet
    if (this.rawLists.length > 0) {
      const ws = XLSX.utils.aoa_to_sheet(this.rawLists);
      XLSX.utils.book_append_sheet(wb, ws, 'Lists');
    }

    // Write file and trigger download
    XLSX.writeFile(wb, 'QNET_2026_Sales_Weekly_Tracker_Exported.xlsx');
  }
}
