import { Controller, Get } from '@nestjs/common';
import { AppService } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {
    this.appService.initialize();
  }

  @Get()
  async get() {
    return await this.appService.getData();
  }
}
