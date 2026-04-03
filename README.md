# Evrak AI

FastAPI tabanli belge yukleme ve AI ile alan cikarma uygulamasi.

## Yerel calistirma

```powershell
python -m venv venv
venv\Scripts\activate
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

## Render ile deploy

1. Kodu GitHub'a gonder.
2. Render'da `New +` -> `Blueprint` sec.
3. Repo'yu bagla.
4. `CLAUDE_API_KEY` ortam degiskenini Render uzerinden ekle.
5. Kalici kullanim icin `DATABASE_URL`, `SUPABASE_URL`, `SUPABASE_ANON_KEY`, `SUPABASE_SECRET_KEY` ortam degiskenlerini de ekle.
6. `SUPABASE_BUCKET=evrak-files` olarak ayarla.
7. Deploy tamamlaninca gelen URL'i telefondan ac.

Render servis tanimi [render.yaml](/C:/evrak-ai/render.yaml) icinde hazir.
