import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelDataService, TrackerActivity, NameCard } from '../../services/excel-data.service';
import { Subscription } from 'rxjs';

@Component({
  selector: 'app-weekly-tracker',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './weekly-tracker.html',
  styleUrl: './weekly-tracker.css'
})
export class WeeklyTrackerComponent implements OnInit, OnDestroy {
  private sub = new Subscription();
  allActivities: TrackerActivity[] = [];
  filteredActivities: TrackerActivity[] = [];
  nameCards: NameCard[] = [];

  // Filter options
  weeks: (string | number)[] = [];
  months: string[] = [];
  names: string[] = [];

  // Filters
  filterWeek = '';
  filterMonth = '';
  filterName = '';

  // Autocomplete search states (Filter toolbar)
  filterNameSearchText = '';
  showFilterNameDropdown = false;
  highlightedFilterNameIndex = -1;

  // Autocomplete search states (Modal form)
  contactSearchText = '';
  showContactDropdown = false;
  highlightedContactIndex = -1;

  // Logging Modals
  showModal = false;
  isEditMode = false;
  modalActivity: TrackerActivity = this.getEmptyActivity();

  // Delete modal
  showDeleteConfirm = false;
  activityToDelete: TrackerActivity | null = null;

  constructor(private dataService: ExcelDataService) {}

  ngOnInit(): void {
    this.sub.add(
      this.dataService.trackerActivities$.subscribe(activities => {
        this.allActivities = activities;
        this.extractFilters(activities);
        this.applyFilters();
      })
    );

    this.sub.add(
      this.dataService.nameCards$.subscribe(cards => {
        this.nameCards = cards;
        this.names = cards.map(c => c.name).filter(n => !!n).sort();
      })
    );
  }

  ngOnDestroy(): void {
    this.sub.unsubscribe();
  }

  private extractFilters(activities: TrackerActivity[]): void {
    const weekSet = new Set<string | number>();
    const monthSet = new Set<string>();

    activities.forEach(a => {
      if (a.weekNo !== undefined && a.weekNo !== '') weekSet.add(a.weekNo);
      if (a.salesMonth) monthSet.add(a.salesMonth.toUpperCase());
    });

    this.weeks = Array.from(weekSet).sort((a, b) => Number(a) - Number(b));
    this.months = Array.from(monthSet).sort();
  }

  public applyFilters(): void {
    const week = this.filterWeek;
    const month = this.filterMonth.toLowerCase();
    const name = this.filterName.toLowerCase();

    this.filteredActivities = this.allActivities.filter(a => {
      const matchesWeek = !week || String(a.weekNo) === String(week);
      const matchesMonth = !month || (a.salesMonth && a.salesMonth.toLowerCase() === month);
      const matchesName = !name || (a.name && a.name.toLowerCase() === name);

      return matchesWeek && matchesMonth && matchesName;
    });

    // Sort activities by date descending
    this.filteredActivities.sort((a, b) => {
      return new Date(b.date).getTime() - new Date(a.date).getTime();
    });
  }

  // Auto-fill attributes when Date changes
  public onDateChange(newDate: string): void {
    if (!newDate) return;
    const dateObj = new Date(newDate);
    if (isNaN(dateObj.getTime())) return;

    // Day of the week
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    this.modalActivity.day = days[dateObj.getDay()];

    // Month
    const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    this.modalActivity.salesMonth = months[dateObj.getMonth()];

    // Rough Week estimation
    const start = new Date(dateObj.getFullYear(), 0, 1);
    const diff = dateObj.getTime() - start.getTime() + ((start.getTimezoneOffset() - dateObj.getTimezoneOffset()) * 60 * 1000);
    const oneDay = 1000 * 60 * 60 * 24;
    const dayOfYear = Math.floor(diff / oneDay);
    this.modalActivity.weekNo = Math.ceil((dayOfYear + start.getDay() + 1) / 7);
  }

  // Getters for filtered contact lists in dropdowns
  get filteredFilterNames(): string[] {
    const s = this.filterNameSearchText.toLowerCase().trim();
    if (!s) return this.names;
    return this.names.filter(n => n.toLowerCase().includes(s));
  }

  get filteredModalNames(): string[] {
    const s = this.contactSearchText.toLowerCase().trim();
    if (!s) return this.names;
    return this.names.filter(n => n.toLowerCase().includes(s));
  }

  // Selection actions for searchable dropdowns
  public onFilterNameSearchInput(): void {
    this.showFilterNameDropdown = true;
    this.highlightedFilterNameIndex = 0;

    if (this.filterName && !this.filteredFilterNames.includes(this.filterName)) {
      this.filterName = '';
      this.applyFilters();
    }
  }

  public onFilterNameKeydown(event: KeyboardEvent): void {
    const items = this.filteredFilterNames;
    if (!this.showFilterNameDropdown || items.length === 0) return;

    if (event.key === 'ArrowDown') {
      event.preventDefault();
      this.highlightedFilterNameIndex = (this.highlightedFilterNameIndex + 1) % items.length;
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      this.highlightedFilterNameIndex = (this.highlightedFilterNameIndex <= 0)
        ? items.length - 1
        : this.highlightedFilterNameIndex - 1;
    } else if (event.key === 'Enter') {
      event.preventDefault();
      const selected = items[this.highlightedFilterNameIndex] || items[0];
      if (selected) {
        this.selectContactForFilter(selected);
      }
    } else if (event.key === 'Escape') {
      this.showFilterNameDropdown = false;
    }
  }

  public selectContactForFilter(name: string): void {
    this.filterName = name;
    this.filterNameSearchText = name;
    this.showFilterNameDropdown = false;
    this.highlightedFilterNameIndex = -1;
    this.applyFilters();
  }

  public clearFilterName(): void {
    this.filterName = '';
    this.filterNameSearchText = '';
    this.showFilterNameDropdown = false;
    this.highlightedFilterNameIndex = -1;
    this.applyFilters();
  }

  public onContactSearchInput(): void {
    this.showContactDropdown = true;
    this.highlightedContactIndex = 0;

    if (this.modalActivity.name && !this.filteredModalNames.includes(this.modalActivity.name)) {
      this.modalActivity.name = '';
    }
  }

  public onContactKeydown(event: KeyboardEvent): void {
    const items = this.filteredModalNames;
    if (!this.showContactDropdown || items.length === 0) return;

    if (event.key === 'ArrowDown') {
      event.preventDefault();
      this.highlightedContactIndex = (this.highlightedContactIndex + 1) % items.length;
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      this.highlightedContactIndex = (this.highlightedContactIndex <= 0)
        ? items.length - 1
        : this.highlightedContactIndex - 1;
    } else if (event.key === 'Enter') {
      event.preventDefault();
      const selected = items[this.highlightedContactIndex] || items[0];
      if (selected) {
        this.selectContactForModal(selected);
      }
    } else if (event.key === 'Escape') {
      this.showContactDropdown = false;
    }
  }

  public selectContactForModal(name: string): void {
    this.modalActivity.name = name;
    this.contactSearchText = name;
    this.showContactDropdown = false;
    this.highlightedContactIndex = -1;
  }

  // CRUD Actions
  public openAddModal(): void {
    this.isEditMode = false;
    this.modalActivity = this.getEmptyActivity();
    this.contactSearchText = '';
    
    // Set default date as today
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    this.modalActivity.date = `${y}-${m}-${d}`;
    this.onDateChange(this.modalActivity.date);

    this.showModal = true;
  }

  public openEditModal(act: TrackerActivity): void {
    this.isEditMode = true;
    this.modalActivity = { ...act };
    this.contactSearchText = act.name || '';
    this.showModal = true;
  }

  public closeModal(): void {
    this.showModal = false;
  }

  public saveActivity(): void {
    if (!this.modalActivity.name) {
      alert('Contact Name is required! Please select a contact.');
      return;
    }

    if (this.isEditMode) {
      this.dataService.editActivity(this.modalActivity);
    } else {
      this.dataService.addActivity(this.modalActivity);
    }

    this.showModal = false;
  }

  public confirmDelete(act: TrackerActivity): void {
    this.activityToDelete = act;
    this.showDeleteConfirm = true;
  }

  public deleteActivity(): void {
    if (this.activityToDelete) {
      this.dataService.deleteActivity(this.activityToDelete.id);
      this.showDeleteConfirm = false;
      this.activityToDelete = null;
    }
  }

  public cancelDelete(): void {
    this.showDeleteConfirm = false;
    this.activityToDelete = null;
  }

  private getEmptyActivity(): TrackerActivity {
    return {
      id: '',
      date: '',
      day: '',
      salesMonth: '',
      weekNo: '',
      name: '',
      info: '',
      infoDate: '',
      invite: '',
      inviteDate: '',
      plan: 'No',
      newNumber: '',
      newConnection: '',
      newLocation: '',
      comments: '',
      notes: ''
    };
  }
}
