# Excel Import Service

Stateless `.xlsx` parser used by Strapi projects that own their own MinIO and staging database.

```bash
npm install
npm test
npm start
```

Required configuration:

```env
EXCEL_SERVICE_API_KEY=development-secret
EXCEL_SOURCE_ALLOWED_HOSTS=minio
```

Production should use `EXCEL_SERVICE_PROJECT_KEYS_JSON` instead of a shared single API key.

See [`../Excel-Microservice.md`](../Excel-Microservice.md) for architecture, API contract, security, deployment, migration and operational notes.
