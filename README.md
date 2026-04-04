# Evrak AI

FastAPI tabanli belge yukleme ve AI ile alan cikarma uygulamasi.

## Yerel calistirma

### Windows

```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn app:app --reload
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload
```

`.env` dosyasi:

```env
CLAUDE_API_KEY=your_anthropic_key
CLAUDE_MODEL=claude-opus-4-6
DATABASE_URL=sqlite:///data/app.db
SUPABASE_URL=
SUPABASE_ANON_KEY=
SUPABASE_SECRET_KEY=
SUPABASE_BUCKET=evrak-files
```

Not:

- `DATABASE_URL` bos birakilirsa uygulama varsayilan olarak yerel `SQLite` dosyasi kullanir.
- `SUPABASE_*` alanlari doldurulursa yuklenen dosyalar `Supabase Storage` uzerine yazilir.
- Windows'tan macOS'a dosya tasindiginda sanal ortam tekrar kurulmalidir.

## Render ile deploy

1. Kodu GitHub'a gonder.
2. Render'da `New +` -> `Blueprint` sec.
3. Repo'yu bagla.
4. `CLAUDE_API_KEY` ortam degiskenini Render uzerinden ekle.
5. Kalici kullanim icin `DATABASE_URL`, `SUPABASE_URL`, `SUPABASE_ANON_KEY`, `SUPABASE_SECRET_KEY` ortam degiskenlerini de ekle.
6. `SUPABASE_BUCKET=evrak-files` olarak ayarla.
7. Deploy tamamlaninca gelen URL'i telefondan ac.

Render servis tanimi [render.yaml](/Users/murattufan/Desktop/evrak-ai/render.yaml) icinde hazir.
