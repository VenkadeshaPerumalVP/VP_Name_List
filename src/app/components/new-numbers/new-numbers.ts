import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelDataService, NewNumber, NameCard } from '../../services/excel-data.service';
import { Subscription } from 'rxjs';

@Component({
  selector: 'app-new-numbers',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './new-numbers.html',
  styleUrl: './new-numbers.css'
})
export class NewNumbersComponent implements OnInit, OnDestroy {
  private sub = new Subscription();
  allNewNumbers: NewNumber[] = [];
  filteredNewNumbers: NewNumber[] = [];

  // Filters
  searchText = '';
  filterStatus = 'No'; // Default to show pending leads

  // CRUD Modals
  showModal = false;
  isEditMode = false;
  modalNumber: NewNumber = this.getEmptyNumber();

  // Promotion Modal
  showPromoteModal = false;
  promotingNumber: NewNumber | null = null;
  promoteCard: Omit<NameCard, 'id'> = this.getEmptyCard();

  // Delete modal
  showDeleteConfirm = false;
  numToDelete: NewNumber | null = null;

  constructor(private dataService: ExcelDataService) {}

  ngOnInit(): void {
    this.sub.add(
      this.dataService.newNumbers$.subscribe(nums => {
        this.allNewNumbers = nums;
        this.applyFilters();
      })
    );
  }

  ngOnDestroy(): void {
    this.sub.unsubscribe();
  }

  public applyFilters(): void {
    const search = this.searchText.toLowerCase().trim();
    const status = this.filterStatus.toLowerCase();

    this.filteredNewNumbers = this.allNewNumbers.filter(num => {
      const matchesSearch = !search ||
        (num.name && num.name.toLowerCase().includes(search)) ||
        (num.contactNumber && num.contactNumber.toLowerCase().includes(search)) ||
        (num.location && num.location.toLowerCase().includes(search)) ||
        (num.notes && num.notes.toLowerCase().includes(search)) ||
        (num.comments && num.comments.toLowerCase().includes(search));

      const matchesStatus = !status || 
        (status === 'all') || 
        (num.addedToMain && num.addedToMain.toLowerCase() === status);

      return matchesSearch && matchesStatus;
    });

    // Sort by date added descending
    this.filteredNewNumbers.sort((a, b) => {
      return new Date(b.dateAdded).getTime() - new Date(a.dateAdded).getTime();
    });
  }

  // CRUD Actions
  public openAddModal(): void {
    this.isEditMode = false;
    this.modalNumber = this.getEmptyNumber();

    // Default today
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    this.modalNumber.dateAdded = `${y}-${m}-${d}`;
    this.modalNumber.salesMonth = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'][today.getMonth()];
    
    // Simple calendar week estimation
    const start = new Date(today.getFullYear(), 0, 1);
    const diff = today.getTime() - start.getTime();
    const oneDay = 1000 * 60 * 60 * 24;
    const dayOfYear = Math.floor(diff / oneDay);
    this.modalNumber.weekNo = Math.ceil((dayOfYear + start.getDay() + 1) / 7);

    this.showModal = true;
  }

  public openEditModal(num: NewNumber): void {
    this.isEditMode = true;
    this.modalNumber = { ...num };
    this.showModal = true;
  }

  public closeModal(): void {
    this.showModal = false;
  }

  public saveNewNumber(): void {
    if (!this.modalNumber.name && !this.modalNumber.contactNumber) {
      alert('Name or Contact Number is required!');
      return;
    }

    if (this.isEditMode) {
      this.dataService.editNewNumber(this.modalNumber);
    } else {
      this.dataService.addNewNumber(this.modalNumber);
    }

    this.showModal = false;
  }

  // Promotion Flow
  public openPromoteModal(num: NewNumber): void {
    this.promotingNumber = num;
    this.promoteCard = {
      name: num.name || '',
      contactNumber: num.contactNumber || '',
      zone: 'WARM', // default zone
      connection: num.connection || '',
      location: num.location || '',
      info: 'Added from Leads',
      infoDate: num.dateAdded || '',
      invite: '',
      inviteDate: '',
      plan: 'No',
      comments: num.comments || '',
      notes: num.notes || ''
    };
    this.showPromoteModal = true;
  }

  public closePromoteModal(): void {
    this.showPromoteModal = false;
    this.promotingNumber = null;
  }

  public executePromotion(): void {
    if (!this.promotingNumber) return;
    this.dataService.promoteToMain(this.promotingNumber.id, this.promoteCard);
    this.closePromoteModal();
  }

  // Delete Action
  public confirmDelete(num: NewNumber): void {
    this.numToDelete = num;
    this.showDeleteConfirm = true;
  }

  public deleteNewNumber(): void {
    if (this.numToDelete) {
      this.dataService.deleteNewNumber(this.numToDelete.id);
      this.showDeleteConfirm = false;
      this.numToDelete = null;
    }
  }

  public cancelDelete(): void {
    this.showDeleteConfirm = false;
    this.numToDelete = null;
  }

  private getEmptyNumber(): NewNumber {
    return {
      id: '',
      salesMonth: '',
      weekNo: '',
      dateAdded: '',
      name: '',
      contactNumber: '',
      connection: '',
      location: '',
      comments: '',
      notes: '',
      addedToMain: 'No'
    };
  }

  private getEmptyCard(): Omit<NameCard, 'id'> {
    return {
      name: '',
      contactNumber: '',
      zone: 'WARM',
      connection: '',
      location: '',
      info: '',
      infoDate: '',
      invite: '',
      inviteDate: '',
      plan: 'No',
      comments: '',
      notes: ''
    };
  }
}
