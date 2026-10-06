# excelword — Học Excel & Word từ cơ bản đến nâng cao

Website học tập: 51 bài video, 391 công thức Excel, 213 kỹ năng Word, 1.820 câu trắc nghiệm ABCD.

- Frontend tĩnh: `website/` (mở bằng `python -m http.server` trong `website/`)
- API Vercel + MongoDB Atlas: `api/` (`MONGODB_URI`, `MONGODB_DB`, `SEED_KEY`)
- Seed DB: `npm install && npm run seed` hoặc `POST /api/seed` với header `x-seed-key`
- Deploy: import repo lên Vercel (`vercel.json` đã cấu hình sẵn)
