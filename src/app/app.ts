import { Component, signal, OnInit, OnDestroy } from '@angular/core';
import { RouterOutlet, RouterLink, RouterLinkActive } from '@angular/router';
import { CommonModule } from '@angular/common';
import { ExcelDataService } from './services/excel-data.service';
import { Subscription } from 'rxjs';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, RouterOutlet, RouterLink, RouterLinkActive],
  templateUrl: './app.html',
  styleUrl: './app.css'
})
export class App implements OnInit, OnDestroy {
  protected readonly title = signal('QNET Sales & NameCards Tracker');
  
  isDataLoaded = false;
  private sub = new Subscription();

  constructor(public dataService: ExcelDataService) {}

  ngOnInit(): void {
    this.sub.add(
      this.dataService.isLoaded$.subscribe(loaded => {
        this.isDataLoaded = loaded;
      })
    );
  }

  ngOnDestroy(): void {
    this.sub.unsubscribe();
  }

  /**
   * Universal Excel Import trigger
   */
  public triggerFileInput(fileInput: HTMLInputElement): void {
    fileInput.click();
  }

  public onFileImport(event: any): void {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    const reader = new FileReader();
    
    reader.onload = (e: any) => {
      const buffer = e.target.result;
      this.isDataLoaded = false;
      
      // Delay slightly for smooth transitions
      setTimeout(() => {
        this.dataService.parseExcelBuffer(buffer);
        // Clear value to allow re-importing the same file
        event.target.value = '';
      }, 500);
    };

    reader.readAsArrayBuffer(file);
  }

  /**
   * Universal Excel Export trigger
   */
  public onFileExport(): void {
    this.dataService.exportToExcel();
  }
}
