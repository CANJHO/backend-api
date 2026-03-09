import { Global, Module } from "@nestjs/common";
import { ConfigModule, ConfigService } from "@nestjs/config";
import { DataSource } from "typeorm";

@Global()
@Module({
  imports: [ConfigModule],
  providers: [
    {
      provide: DataSource,
      inject: [ConfigService],
      useFactory: async (cfg: ConfigService) => {
        const ds = new DataSource({
          type: "postgres",
          host: cfg.get<string>("DB_HOST"),
          port: parseInt(cfg.get<string>("DB_PORT", "5432")),
          username: cfg.get<string>("DB_USER"),
          password: cfg.get<string>("DB_PASS"),
          database: cfg.get<string>("DB_NAME"),
        });

        if (!ds.isInitialized) {
          await ds.initialize();
        }

        return ds;
      },
    },
  ],
  exports: [DataSource],
})
export class DatabaseModule {}
