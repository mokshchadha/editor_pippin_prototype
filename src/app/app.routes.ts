import { Routes } from '@angular/router';
import { CommitmentCollationComponent } from './commitment-collation/commitment-collation.component';

export const routes: Routes = [
    { path: 'commitment-collation', component: CommitmentCollationComponent },
    { path: '', redirectTo: 'commitment-collation', pathMatch: 'full' }
];
