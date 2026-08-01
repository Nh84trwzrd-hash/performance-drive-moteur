"""
Moteur de génération du fichier "Performance Drive" à partir du fichier
Preparation hebdomadaire. Expose un unique endpoint HTTP que le workflow
n8n appelle (noeud "Generer fichier Performance Drive").

Contrat de l'API :

POST /generate?taux_actuelle=0.0245&taux_precedente=0.0219
Corps de la requête : le fichier Preparation .xlsx brut (octets bruts, pas
de multipart).
Réponse : le fichier "Performance Drive S{semaine}.xlsx" généré, en pièce
jointe (Content-Disposition), semaine détectée automatiquement à partir de
la date trouvée dans le fichier source.

Lancement local :
    pip install -r requirements.txt
    uvicorn main:app --reload

Déploiement (Render/Railway/Fly.io) :
    Commande de démarrage : uvicorn main:app --host 0.0.0.0 --port $PORT
"""
import os
import tempfile
import shutil

from fastapi import FastAPI, Request, Query, HTTPException
from fastapi.responses import FileResponse, JSONResponse

import gen_lib

app = FastAPI(title="Moteur Performance Drive")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate")
async def generate(
    request: Request,
    taux_actuelle: float | None = Query(default=None, description="Taux de rupture de la semaine courante (fraction, ex: 0.0245 pour 2.45%)"),
    taux_precedente: float | None = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: int | None = Query(default=None, description="Numéro de semaine indiqué par l'expéditeur (informatif — la semaine réellement utilisée est déduite de la date dans le fichier)"),
):
    body = await request.body()
    if not body:
        raise HTTPException(status_code=400, detail="Corps de la requête vide : le fichier Preparation .xlsx est attendu en octets bruts.")

    work_dir = tempfile.mkdtemp(prefix="perfdrive_")
    try:
        src_path = os.path.join(work_dir, "source.xlsx")
        with open(src_path, "wb") as f:
            f.write(body)

        try:
            detected_week, employees = gen_lib.get_week_and_employees(src_path)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Impossible de lire le fichier Preparation envoyé : {e}")

        if not employees:
            raise HTTPException(status_code=400, detail="Aucun employé détecté dans l'onglet Preparation — vérifie le fichier source.")

        if semaine is not None and int(semaine) != detected_week:
            # On fait confiance à la date réelle du fichier, pas au numéro indiqué
            # dans l'email (qui peut être une erreur de saisie). On le signale
            # juste dans les logs du service.
            print(f"Avertissement: semaine indiquée dans l'email ({semaine}) != semaine déduite du fichier ({detected_week}). Semaine déduite utilisée.")

        out_name = f"Performance Drive S{detected_week}.xlsx"
        out_path = os.path.join(work_dir, out_name)

        gen_lib.build(
            src_path,
            out_path,
            taux_actuelle=taux_actuelle,
            taux_precedente=taux_precedente,
        )

        return FileResponse(
            out_path,
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération du fichier : {e}")
    # Note: work_dir is intentionally not cleaned up synchronously here since
    # FileResponse streams the file after this function returns. The OS temp
    # directory is periodically cleaned by the host; for high volume, add a
    # background task to remove work_dir after the response is sent.
