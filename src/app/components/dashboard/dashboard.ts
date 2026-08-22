import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { ExcelDataService, NameCard, TrackerActivity } from '../../services/excel-data.service';
import { Subscription } from 'rxjs';

interface WeekActivity {
  week: string;
  info: number;
  invite: number;
  plan: number;
  total: number;
}

interface ZoneCount {
  zone: string;
  count: number;
  percentage: number;
  color: string;
}

@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.css'
})
export class DashboardComponent implements OnInit, OnDestroy {
  private sub = new Subscription();

  totalContacts = 0;
  infoUpdatesCount = 0;
  inviteUpdatesCount = 0;
  planUpdatesCount = 0;
  newNumbersCount = 0;

  zoneCounts: ZoneCount[] = [];
  weeklyActivities: WeekActivity[] = [];
  topLocations: { name: string; count: number; percentage: number }[] = [];
  topConnections: { name: string; count: number; percentage: number }[] = [];

  // SVG Chart states
  chartWidth = 500;
  chartHeight = 220;
  pieRadius = 70;
  pieCX = 100;
  pieCY = 100;

  constructor(private dataService: ExcelDataService) {}

  ngOnInit(): void {
    this.sub.add(
      this.dataService.nameCards$.subscribe(cards => {
        this.calculateCardStats(cards);
      })
    );

    this.sub.add(
      this.dataService.trackerActivities$.subscribe(activities => {
        this.calculateTrackerStats(activities);
      })
    );

    this.sub.add(
      this.dataService.newNumbers$.subscribe(nums => {
        this.newNumbersCount = nums.filter(n => n.addedToMain !== 'Yes').length;
      })
    );
  }

  ngOnDestroy(): void {
    this.sub.unsubscribe();
  }

  private calculateCardStats(cards: NameCard[]): void {
    this.totalContacts = cards.length;
    
    // Count zone distributions
    const zonesMap = new Map<string, number>();
    const locationsMap = new Map<string, number>();
    const connectionsMap = new Map<string, number>();
    
    let planCount = 0;
    
    cards.forEach(card => {
      const z = card.zone || 'Unknown';
      zonesMap.set(z, (zonesMap.get(z) || 0) + 1);
      
      const loc = card.location || 'Unknown';
      locationsMap.set(loc, (locationsMap.get(loc) || 0) + 1);

      const conn = card.connection || 'Unknown';
      connectionsMap.set(conn, (connectionsMap.get(conn) || 0) + 1);
      
      if (card.plan && card.plan.trim() !== '' && card.plan.toLowerCase() !== 'no') {
        planCount++;
      }
    });

    this.planUpdatesCount = planCount;

    // Format Zones
    const colorsMap: Record<string, string> = {
      'HOT': 'var(--color-hot-from)',
      'WARM': 'var(--color-warm-from)',
      'COLD': 'var(--color-cold-from)',
      'Unknown': 'var(--text-muted)'
    };
    
    let totalZones = 0;
    zonesMap.forEach(count => totalZones += count);
    
    this.zoneCounts = [];
    zonesMap.forEach((count, zone) => {
      this.zoneCounts.push({
        zone,
        count,
        percentage: totalZones > 0 ? Math.round((count / totalZones) * 100) : 0,
        color: colorsMap[zone.toUpperCase()] || 'var(--accent-violet)'
      });
    });
    this.zoneCounts.sort((a, b) => b.count - a.count);

    // Format top Locations
    this.topLocations = Array.from(locationsMap.entries())
      .map(([name, count]) => ({
        name,
        count,
        percentage: this.totalContacts > 0 ? Math.round((count / this.totalContacts) * 100) : 0
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    // Format top Connections
    this.topConnections = Array.from(connectionsMap.entries())
      .map(([name, count]) => ({
        name,
        count,
        percentage: this.totalContacts > 0 ? Math.round((count / this.totalContacts) * 100) : 0
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);
  }

  private calculateTrackerStats(activities: TrackerActivity[]): void {
    // Info updates & Invite updates count
    this.infoUpdatesCount = activities.filter(a => a.info && a.info.trim() !== '' && a.info.toLowerCase() !== 'no').length;
    this.inviteUpdatesCount = activities.filter(a => a.invite && a.invite.trim() !== '' && a.invite.toLowerCase() !== 'no').length;

    // Group activities by week
    const weeksMap = new Map<string, { info: number; invite: number; plan: number }>();

    activities.forEach(act => {
      const weekLabel = act.weekNo ? `Week ${act.weekNo}` : 'Unknown';
      const stats = weeksMap.get(weekLabel) || { info: 0, invite: 0, plan: 0 };
      
      if (act.info && act.info.trim() !== '' && act.info.toLowerCase() !== 'no') stats.info++;
      if (act.invite && act.invite.trim() !== '' && act.invite.toLowerCase() !== 'no') stats.invite++;
      if (act.plan && act.plan.trim() !== '' && act.plan.toLowerCase() !== 'no') stats.plan++;

      weeksMap.set(weekLabel, stats);
    });

    this.weeklyActivities = Array.from(weeksMap.entries())
      .map(([week, stats]) => ({
        week,
        info: stats.info,
        invite: stats.invite,
        plan: stats.plan,
        total: stats.info + stats.invite + stats.plan
      }))
      .filter(w => w.week !== 'Unknown') // filter out empty labels
      .sort((a, b) => {
        const numA = parseInt(a.week.replace(/\D/g, ''), 10) || 0;
        const numB = parseInt(b.week.replace(/\D/g, ''), 10) || 0;
        return numA - numB;
      })
      .slice(-8); // Show last 8 weeks
  }

  // --- SVG Pie Chart Drawing Helpers ---
  getPieSlices(): { d: string; color: string; label: string; percentage: number }[] {
    let accumulatedAngle = 0;
    const slices: { d: string; color: string; label: string; percentage: number }[] = [];
    
    this.zoneCounts.forEach(zone => {
      if (zone.percentage === 0) return;
      const angle = (zone.percentage / 100) * 360;
      
      // Calculate coordinates
      const x1 = this.pieCX + this.pieRadius * Math.sin((accumulatedAngle * Math.PI) / 180);
      const y1 = this.pieCY - this.pieRadius * Math.cos((accumulatedAngle * Math.PI) / 180);
      
      accumulatedAngle += angle;
      
      const x2 = this.pieCX + this.pieRadius * Math.sin((accumulatedAngle * Math.PI) / 180);
      const y2 = this.pieCY - this.pieRadius * Math.cos((accumulatedAngle * Math.PI) / 180);
      
      const largeArc = angle > 180 ? 1 : 0;
      const d = `M ${this.pieCX} ${this.pieCY} L ${x1} ${y1} A ${this.pieRadius} ${this.pieRadius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
      
      slices.push({
        d,
        color: zone.color,
        label: zone.zone,
        percentage: zone.percentage
      });
    });

    return slices;
  }

  // --- SVG Bar Chart Drawing Helpers ---
  getBarChartMax(): number {
    if (this.weeklyActivities.length === 0) return 10;
    const maxVal = Math.max(...this.weeklyActivities.map(w => Math.max(w.info, w.invite, w.plan)));
    return maxVal > 0 ? maxVal * 1.15 : 10; // add padding
  }

  getBarHeight(value: number, max: number): number {
    const plotHeight = this.chartHeight - 40; // reserve space for text
    return (value / max) * plotHeight;
  }
}
