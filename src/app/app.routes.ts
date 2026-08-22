import { Routes } from '@angular/router';
import { DashboardComponent } from './components/dashboard/dashboard';
import { NameCardsComponent } from './components/name-cards/name-cards';
import { WeeklyTrackerComponent } from './components/weekly-tracker/weekly-tracker';
import { NewNumbersComponent } from './components/new-numbers/new-numbers';

export const routes: Routes = [
  { path: 'dashboard', component: DashboardComponent },
  { path: 'name-cards', component: NameCardsComponent },
  { path: 'weekly-tracker', component: WeeklyTrackerComponent },
  { path: 'new-numbers', component: NewNumbersComponent },
  { path: '', redirectTo: 'dashboard', pathMatch: 'full' },
  { path: '**', redirectTo: 'dashboard' }
];

