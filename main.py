"""
Moteur de génération du fichier "Performance Drive" à partir du fichier
Preparation hebdomadaire et (optionnellement) d'un planning PDF. Expose un
unique endpoint HTTP que le workflow n8n appelle (noeud "Generer fichier
Performance Drive").

Contrat de l'API :

POST /generate?taux_actuelle=0.0245&taux_precedente=0.0219&productivite_s1={"MONTESCOT1":12.5}
Corps de la requête : multipart/form-data avec deux champs fichier :
  - preparation : le fichier Preparation .xlsx (obligatoire)
  - planning    : le planning PDF hebdomadaire (optionnel — si absent, la
                  colonne "Heure travaillée" reste manuelle comme avant)
Réponse : le fichier "Performance Drive S{semaine}.xlsx" généré, en pièce
jointe (Content-Disposition), semaine détectée automatiquement à partir de
la date trouvée dans le fichier source. Un header "X-Employee-Productivity"
contient un JSON {matricule: {nom, productivite_h}} calculé cette semaine,
à faire persister par n8n pour servir de S-1 la semaine suivante.

Lancement local :
    pip install -r requirements.txt
    uvicorn main:app --reload

Déploiement (Render/Railway/Fly.io) :
    Commande de démarrage : uvicorn main:app --host 0.0.0.0 --port $PORT
"""
import os
import json
import tempfile

from fastapi import FastAPI, File, UploadFile, Query, HTTPException
from fastapi.responses import FileResponse
from typing import Optional

import gen_lib

app = FastAPI(title="Moteur Performance Drive")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generate")
async def generate(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None, description="Taux de rupture de la semaine courante (fraction, ex: 0.0245 pour 2.45%)"),
    taux_precedente: Optional[float] = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: Optional[int] = Query(default=None, description="Numéro de semaine indiqué par l'expéditeur (informatif — la semaine réellement utilisée est déduite de la date dans le fichier)"),
    productivite_s1: Optional[str] = Query(default=None, description="JSON {matricule: productivite_h} de la semaine précédente, pour remplir automatiquement la colonne S-1"),
):
    prep_bytes = await preparation.read()
    if not prep_bytes:
        raise HTTPException(status_code=400, detail="Fichier Preparation vide ou manquant.")

    productivite_s1_map = {}
    if productivite_s1:
        try:
            parsed = json.loads(productivite_s1)
            # accepte soit {matricule: valeur}, soit {matricule: {"productivite_h": valeur}}
            for k, v in parsed.items():
                if isinstance(v, dict):
                    productivite_s1_map[k] = v.get('productivite_h')
                else:
                    productivite_s1_map[k] = v
        except Exception as e:
            print(f"Avertissement: productivite_s1 illisible ({e}), ignore.")

    work_dir = tempfile.mkdtemp(prefix="perfdrive_")
    try:
        src_path = os.path.join(work_dir, "source.xlsx")
        with open(src_path, "wb") as f:
            f.write(prep_bytes)

        planning_path = None
        if planning is not None:
            planning_bytes = await planning.read()
            if planning_bytes:
                planning_path = os.path.join(work_dir, "planning.pdf")
                with open(planning_path, "wb") as f:
                    f.write(planning_bytes)

        try:
            detected_week, employees = gen_lib.get_week_and_employees(src_path)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Impossible de lire le fichier Preparation envoyé : {e}")

        if not employees:
            raise HTTPException(status_code=400, detail="Aucun employé détecté dans l'onglet Preparation — vérifie le fichier source.")

        if semaine is not None and int(semaine) != detected_week:
            print(f"Avertissement: semaine indiquée dans l'email ({semaine}) != semaine déduite du fichier ({detected_week}). Semaine déduite utilisée.")

        out_name = f"Performance Drive S{detected_week}.xlsx"
        out_path = os.path.join(work_dir, out_name)

        try:
            week, n, employee_productivity = gen_lib.build(
                src_path,
                out_path,
                taux_actuelle=taux_actuelle,
                taux_precedente=taux_precedente,
                planning_path=planning_path,
                productivite_s1=productivite_s1_map,
            )
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Erreur lors de la lecture du planning ou de la génération : {e}")

        headers = {
            "Content-Disposition": f'attachment; filename="{out_name}"',
            "X-Detected-Week": str(week),
            "X-Employee-Productivity": json.dumps(employee_productivity, ensure_ascii=True),
        }

        return FileResponse(
            out_path,
            filename=out_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération du fichier : {e}")
    # Note: work_dir is intentionally not cleaned up synchronously here since
    # FileResponse streams the file after this function returns. The OS temp
    # directory is periodically cleaned by the host; for high volume, add a
    # background task to remove work_dir after the response is sent.
