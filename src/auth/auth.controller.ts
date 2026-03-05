import { Controller, Post, Body } from '@nestjs/common';
import { Public } from '../common/public.decorator';
import { AuthService } from './auth.service';

@Controller('auth')
export class AuthController {
  constructor(private auth: AuthService) {}

  @Public() // 👈 IMPORTANTE: ESTA LÍNEA HACE QUE LOGIN NO REQUIERA TOKEN
  @Post('login')
  login(@Body() dto: { documento: string; password: string }) {
    return this.auth.login(dto.documento, dto.password);
  }
}
