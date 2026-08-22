import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { ExcelDataService, NameCard } from '../../services/excel-data.service';
import { Subscription } from 'rxjs';

@Component({
  selector: 'app-name-cards',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './name-cards.html',
  styleUrl: './name-cards.css'
})
export class NameCardsComponent implements OnInit, OnDestroy {
  private sub = new Subscription();
  allCards: NameCard[] = [];
  filteredCards: NameCard[] = [];
  paginatedCards: NameCard[] = [];

  // Search & Filters
  searchText = '';
  selectedZone = '';
  selectedConnection = '';
  zones: string[] = ['HOT', 'WARM', 'COLD'];
  connections: string[] = [];

  // Pagination
  currentPage = 1;
  pageSize = 24;
  totalPages = 1;

  // Add / Edit Modal
  showModal = false;
  isEditMode = false;
  modalCard: NameCard = this.getEmptyCard();

  // Delete Modal
  showDeleteConfirm = false;
  cardToDelete: NameCard | null = null;

  constructor(private dataService: ExcelDataService) {}

  ngOnInit(): void {
    this.sub.add(
      this.dataService.nameCards$.subscribe(cards => {
        this.allCards = cards;
        this.extractFilterOptions(cards);
        this.applyFilters();
      })
    );
  }

  ngOnDestroy(): void {
    this.sub.unsubscribe();
  }

  private extractFilterOptions(cards: NameCard[]): void {
    // Dynamic connection types
    const connSet = new Set<string>();
    cards.forEach(c => {
      if (c.connection) connSet.add(c.connection);
    });
    this.connections = Array.from(connSet).sort();
    
    // Zone extraction (if any other than HOT/WARM/COLD)
    const zoneSet = new Set<string>(['HOT', 'WARM', 'COLD']);
    cards.forEach(c => {
      if (c.zone) zoneSet.add(c.zone.toUpperCase());
    });
    this.zones = Array.from(zoneSet).filter(z => z !== '').sort();
  }

  public applyFilters(): void {
    const search = this.searchText.toLowerCase().trim();
    const zone = this.selectedZone.toLowerCase();
    const conn = this.selectedConnection.toLowerCase();

    this.filteredCards = this.allCards.filter(card => {
      const matchesSearch = !search ||
        (card.name && card.name.toLowerCase().includes(search)) ||
        (card.contactNumber && card.contactNumber.toLowerCase().includes(search)) ||
        (card.location && card.location.toLowerCase().includes(search)) ||
        (card.info && card.info.toLowerCase().includes(search)) ||
        (card.invite && card.invite.toLowerCase().includes(search)) ||
        (card.notes && card.notes.toLowerCase().includes(search)) ||
        (card.comments && card.comments.toLowerCase().includes(search));

      const matchesZone = !zone || (card.zone && card.zone.toLowerCase() === zone);
      const matchesConnection = !conn || (card.connection && card.connection.toLowerCase() === conn);

      return matchesSearch && matchesZone && matchesConnection;
    });

    this.totalPages = Math.ceil(this.filteredCards.length / this.pageSize) || 1;
    if (this.currentPage > this.totalPages) {
      this.currentPage = 1;
    }
    this.updatePaginatedCards();
  }

  private updatePaginatedCards(): void {
    const start = (this.currentPage - 1) * this.pageSize;
    const end = start + this.pageSize;
    this.paginatedCards = this.filteredCards.slice(start, end);
  }

  // Pagination Actions
  public prevPage(): void {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.updatePaginatedCards();
      this.scrollToTop();
    }
  }

  public nextPage(): void {
    if (this.currentPage < this.totalPages) {
      this.currentPage++;
      this.updatePaginatedCards();
      this.scrollToTop();
    }
  }

  public setPage(page: number): void {
    if (page >= 1 && page <= this.totalPages) {
      this.currentPage = page;
      this.updatePaginatedCards();
      this.scrollToTop();
    }
  }

  public getPages(): number[] {
    const pages: number[] = [];
    const maxVisible = 5;
    let start = Math.max(1, this.currentPage - 2);
    let end = Math.min(this.totalPages, start + maxVisible - 1);
    
    if (end - start < maxVisible - 1) {
      start = Math.max(1, end - maxVisible + 1);
    }
    
    for (let i = start; i <= end; i++) {
      pages.push(i);
    }
    return pages;
  }

  private scrollToTop(): void {
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // CRUD Trigger Actions
  public openAddModal(): void {
    this.isEditMode = false;
    this.modalCard = this.getEmptyCard();
    this.showModal = true;
  }

  public openEditModal(card: NameCard): void {
    this.isEditMode = true;
    this.modalCard = { ...card }; // Deep copy
    this.showModal = true;
  }

  public closeModal(): void {
    this.showModal = false;
  }

  public saveCard(): void {
    if (!this.modalCard.name.trim()) {
      alert('Name is required!');
      return;
    }
    
    if (this.isEditMode) {
      this.dataService.editNameCard(this.modalCard);
    } else {
      this.dataService.addNameCard(this.modalCard);
    }
    
    this.showModal = false;
  }

  public confirmDelete(card: NameCard): void {
    this.cardToDelete = card;
    this.showDeleteConfirm = true;
  }

  public deleteCard(): void {
    if (this.cardToDelete) {
      this.dataService.deleteNameCard(this.cardToDelete.id);
      this.showDeleteConfirm = false;
      this.cardToDelete = null;
    }
  }

  public cancelDelete(): void {
    this.showDeleteConfirm = false;
    this.cardToDelete = null;
  }

  // Helper getters/converters
  public getAvatarInitials(name: string): string {
    if (!name) return '??';
    const parts = name.trim().split(/\s+/);
    if (parts.length >= 2) {
      return (parts[0][0] + parts[1][0]).toUpperCase();
    }
    return name.slice(0, 2).toUpperCase();
  }

  public getAvatarGradient(name: string): string {
    const gradients = [
      'linear-gradient(135deg, #ec4899, #8b5cf6)',
      'linear-gradient(135deg, #3b82f6, #06b6d4)',
      'linear-gradient(135deg, #10b981, #3b82f6)',
      'linear-gradient(135deg, #f59e0b, #ec4899)',
      'linear-gradient(135deg, #8b5cf6, #d946ef)'
    ];
    let hash = 0;
    for (let i = 0; i < name.length; i++) {
      hash = name.charCodeAt(i) + ((hash << 5) - hash);
    }
    const index = Math.abs(hash) % gradients.length;
    return gradients[index];
  }

  private getEmptyCard(): NameCard {
    return {
      id: '',
      name: '',
      contactNumber: '',
      zone: 'WARM',
      connection: '',
      location: '',
      info: '',
      infoDate: '',
      invite: '',
      inviteDate: '',
      plan: '',
      comments: '',
      notes: ''
    };
  }
}
