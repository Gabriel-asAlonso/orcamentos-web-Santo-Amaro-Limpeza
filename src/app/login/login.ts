import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { FormsModule } from '@angular/forms';
import { Router, RouterLink } from '@angular/router';
import {
  ArrowRight,
  Eye,
  EyeOff,
  Lock,
  Mail,
  ShieldCheck,
  LucideAngularModule
} from 'lucide-angular';

@Component({
  selector: 'app-login',
  imports: [
    CommonModule,
    FormsModule,
    RouterLink,
    LucideAngularModule
  ],
  templateUrl: './login.html',
  styleUrl: './login.scss',
})
export class Login {

  readonly ArrowRight = ArrowRight;
  readonly Eye = Eye;
  readonly EyeOff = EyeOff;
  readonly Lock = Lock;
  readonly Mail = Mail;
  readonly ShieldCheck = ShieldCheck;
  public currentYear = new Date().getFullYear();

  email = '';
  password = '';
  showPassword = false;
  remember = true;
  loading = false;
  errorMessage = '';

  logoUrl = '/logoQuadrada.jpg';
  private readonly apiUrl = 'http://localhost:3000';

  constructor(
    private router: Router,
    private http: HttpClient
  ) {}

  handleSubmit(): void {
    this.errorMessage = '';

    const email = this.email.trim();
    const password = this.password.trim();

    if (!email || !password) {
      this.errorMessage = 'Informe email e senha para continuar.';
      return;
    }

    this.loading = true;

    this.http.post<LoginResponse>(`${this.apiUrl}/auth/login`, { email, password }).subscribe({
      next: (response) => {
        const storage = this.remember ? localStorage : sessionStorage;

        storage.setItem('authToken', response.token);
        storage.setItem('authUser', JSON.stringify(response.user));

        if (this.remember) {
          sessionStorage.removeItem('authToken');
          sessionStorage.removeItem('authUser');
        } else {
          localStorage.removeItem('authToken');
          localStorage.removeItem('authUser');
        }

        this.loading = false;
        this.router.navigate(['/orcamentos']);
      },
      error: (error: HttpErrorResponse) => {
        this.loading = false;
        this.errorMessage = error.error?.message ?? 'Não foi possível entrar. Tente novamente.';
      }
    });
  }

  togglePassword(): void {
    this.showPassword = !this.showPassword;
  }

}

interface LoginResponse {
  message: string;
  token: string;
  user: {
    id: string;
    name: string;
    email: string;
  };
}
