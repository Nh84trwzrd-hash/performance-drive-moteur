"""
Moteur de génération du fichier "Performance Drive" à partir du fichier
Preparation hebdomadaire et (optionnellement) d'un planning PDF. Expose les
endpoints HTTP que le workflow n8n appelle.

Contrat de l'API :

POST /generate?taux_actuelle=0.0245&taux_precedente=0.0219&productivite_s1={"MONTESCOT1":12.5}
Corps de la requête : multipart/form-data avec deux champs fichier :
  - preparation : le fichier Preparation .xlsx (obligatoire)
  - planning    : le planning PDF hebdomadaire (optionnel — si absent, la
                  colonne "Heure travaillée" reste manuelle comme avant)
Réponse : le fichier "Performance Drive S{semaine}.xlsx" généré, en pièce
jointe (Content-Disposition).

POST /productivity : mêmes paramètres que /generate, mais renvoie un JSON
simple (pas de fichier) avec la semaine détectée et la productivité calculée
par employé cette semaine : {"detected_week": 31, "employee_productivity":
{"MONTESCOT1": {"nom": "...", "productivite_h": 54.1}, ...}}. Ce second appel
existe pour que n8n puisse persister l'historique par employé sans avoir à
lire des en-têtes HTTP personnalisés sur une réponse fichier.

Lancement local :
    pip install -r requirements.txt
    uvicorn main:app --reload

Déploiement (Render/Railway/Fly.io) :
    Commande de démarrage : uvicorn main:app --host 0.0.0.0 --port $PORT
"""
import os
import json
import tempfile
import base64

from fastapi import FastAPI, File, UploadFile, Query, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from typing import Optional

import gen_lib

app = FastAPI(title="Moteur Performance Drive")


def _parse_name_aliases(name_aliases: Optional[str]) -> dict:
    """Parse le parametre `name_aliases` (JSON, liste de paires
    [nom_preparation, nom_planning]) envoye par n8n -- ces paires viennent
    d'une Data Table persistante (employee_name_aliases), pas du code : un
    alias confirme par l'utilisateur est ajoute comme une ligne de donnee,
    relue a chaque generation, sans avoir besoin de redeployer le service.
    C'est le mecanisme qui permet au systeme d'apprendre de ses erreurs de
    correlation planning/preparation au fil des semaines."""
    if not name_aliases:
        return {}
    try:
        raw_pairs = json.loads(name_aliases)
        return gen_lib.parse_alias_pairs(raw_pairs)
    except Exception as e:
        print(f"Avertissement: name_aliases illisible ({e}), ignore.")
        return {}


@app.get("/health")
def health():
    return {"status": "ok"}


def _parse_historique_complet(historique_complet: Optional[str]) -> dict:
    """Parse le parametre `historique_complet` (JSON {matricule: {semaine:
    productivite_h}}) envoye par n8n -- construit a partir de TOUTES les
    lignes de la Data Table productivite_employes_historique (pas juste la
    semaine precedente comme productivite_s1). Sert a tracer la courbe
    d'evolution multi-semaines par employe dans gen_lib.build(). Un parametre
    absent/illisible desactive simplement la courbe, sans erreur."""
    if not historique_complet:
        return {}
    try:
        parsed = json.loads(historique_complet)
        if not isinstance(parsed, dict):
            return {}
        return parsed
    except Exception as e:
        print(f"Avertissement: historique_complet illisible ({e}), ignore (pas de courbe d'evolution).")
        return {}


def _parse_productivite_s1(productivite_s1: Optional[str]) -> dict:
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
    return productivite_s1_map


async def _run_build(preparation: UploadFile, planning: Optional[UploadFile],
                      taux_actuelle, taux_precedente, semaine, productivite_s1, name_aliases=None,
                      historique_complet=None):
    prep_bytes = await preparation.read()
    if not prep_bytes:
        raise HTTPException(status_code=400, detail="Fichier Preparation vide ou manquant.")

    productivite_s1_map = _parse_productivite_s1(productivite_s1)
    name_aliases_map = _parse_name_aliases(name_aliases)
    historique_complet_map = _parse_historique_complet(historique_complet)

    work_dir = tempfile.mkdtemp(prefix="perfdrive_")
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
            name_aliases=name_aliases_map,
            historique_complet=historique_complet_map,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la lecture du planning ou de la génération : {e}")

    return work_dir, out_path, out_name, week, employee_productivity


@app.post("/generate")
async def generate(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None, description="Taux de rupture de la semaine courante (fraction, ex: 0.0245 pour 2.45%)"),
    taux_precedente: Optional[float] = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: Optional[int] = Query(default=None, description="Numéro de semaine indiqué par l'expéditeur (informatif)"),
    productivite_s1: Optional[str] = Query(default=None, description="JSON {matricule: productivite_h} de la semaine précédente"),
    historique_complet: Optional[str] = Query(default=None, description='JSON {matricule: {semaine: productivite_h}} de TOUTES les semaines passees connues, pour la courbe d\'evolution multi-semaines'),
):
    try:
        work_dir, out_path, out_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1,
            historique_complet=historique_complet,
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
    # FileResponse streams the file after this function returns.


@app.post("/productivity")
async def productivity(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None),
    taux_precedente: Optional[float] = Query(default=None),
    semaine: Optional[int] = Query(default=None),
    productivite_s1: Optional[str] = Query(default=None),
):
    """Meme traitement que /generate mais renvoie uniquement un JSON (semaine
    detectee + productivite par employe), sans le fichier. Appele par n8n en
    parallele de /generate pour alimenter l'historique par employe."""
    try:
        _work_dir, _out_path, _out_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1,
        )
        return JSONResponse({
            "detected_week": week,
            "employee_productivity": employee_productivity,
        })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors du calcul de productivité : {e}")


@app.post("/podium")
async def podium(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None, description="Taux de rupture de la semaine courante (fraction)"),
    taux_precedente: Optional[float] = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: Optional[int] = Query(default=None),
    productivite_s1: Optional[str] = Query(default=None),
):
    """Meme traitement que /generate mais renvoie le PDF "podium" visuel
    (classement + rappel taux de rupture, charte Intermarché) destine a etre
    affiche en salle de pause. Appele par n8n en parallele de /generate."""
    try:
        work_dir, _out_path, _out_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1,
        )
        pdf_name = f"Podium Performance Drive S{week}.pdf"
        pdf_path = os.path.join(work_dir, pdf_name)
        gen_lib.generate_podium_pdf(
            pdf_path, week, employee_productivity,
            taux_actuelle=taux_actuelle, taux_precedente=taux_precedente,
        )
        return FileResponse(
            pdf_path,
            filename=pdf_name,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{pdf_name}"'},
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération du podium : {e}")


@app.post("/generate_package")
async def generate_package(
    preparation: UploadFile = File(..., description="Fichier Preparation .xlsx"),
    planning: Optional[UploadFile] = File(default=None, description="Planning PDF hebdomadaire (optionnel)"),
    taux_actuelle: Optional[float] = Query(default=None, description="Taux de rupture de la semaine courante (fraction)"),
    taux_precedente: Optional[float] = Query(default=None, description="Taux de rupture de la semaine précédente (fraction)"),
    semaine: Optional[int] = Query(default=None),
    productivite_s1: Optional[str] = Query(default=None, description="JSON {matricule: productivite_h} de la semaine précédente"),
    name_aliases: Optional[str] = Query(default=None, description='JSON [[nom_preparation, nom_planning], ...] issu de la Data Table employee_name_aliases'),
    skip_ai_review: Optional[bool] = Query(default=False, description="Si vrai, saute la relecture visuelle IA (utile pour tester rapidement)"),
    historique_complet: Optional[str] = Query(default=None, description='JSON {matricule: {semaine: productivite_h}} de TOUTES les semaines passees connues, pour la courbe d\'evolution multi-semaines'),
):
    """Point d'entree unique pour le pipeline "verifie avant d'envoyer" :
    genere le xlsx ET le podium PDF, les fait recontroler independamment
    (heures recoupees avec le planning, formules dynamiques, mise en page,
    relecture visuelle IA), puis renvoie un rapport JSON avec les deux
    fichiers en base64.

    - status="ok" : aucune erreur bloquante detectee. `warnings` peut
      contenir des points a verifier (ex: employe non retrouve dans le
      planning) qui ne doivent PAS bloquer l'envoi -- n8n les inclut dans
      une notification separee, en parallele de l'envoi normal.
    - status="blocked" : au moins une erreur bloquante (heures incoherentes
      avec le planning recalcule, formule figee, formule cassee, mise en
      page cassee). Le fichier est quand meme renvoye en base64 (pour
      relecture manuelle) mais n8n ne doit ni l'envoyer par email ni
      l'archiver sur Drive tant que ce n'est pas resolu.

    Les problemes releves par la relecture visuelle IA sont toujours classes
    en warnings (jamais bloquants) : c'est un controle moins fiable/plus
    subjectif que les controles deterministes, on ne veut pas qu'un faux
    positif de l'IA bloque la cadence hebdomadaire a lui seul.
    """
    try:
        work_dir, xlsx_path, xlsx_name, week, employee_productivity = await _run_build(
            preparation, planning, taux_actuelle, taux_precedente, semaine, productivite_s1, name_aliases,
            historique_complet=historique_complet,
        )

        pdf_name = f"Podium Performance Drive S{week}.pdf"
        pdf_path = os.path.join(work_dir, pdf_name)
        podium_error = None
        try:
            gen_lib.generate_podium_pdf(
                pdf_path, week, employee_productivity,
                taux_actuelle=taux_actuelle, taux_precedente=taux_precedente,
            )
        except Exception as e:
            # Le generateur leve volontairement une erreur explicite si la
            # mise en page deborderait (voir gen_lib.generate_podium_pdf) --
            # c'est une erreur bloquante, pas un crash a 500 : on la
            # remonte dans le rapport comme les autres.
            podium_error = f"Génération du podium échouée : {e}"

        errors, warnings = [], []
        if podium_error:
            errors.append(podium_error)

        # Reparse independant du planning + verification des formules /
        # coherence des heures ecrites dans le fichier genere.
        prep_path = os.path.join(work_dir, "source.xlsx")
        planning_path = os.path.join(work_dir, "planning.pdf")
        planning_path = planning_path if os.path.exists(planning_path) else None
        name_aliases_map = _parse_name_aliases(name_aliases)

        xlsx_report = gen_lib.verify_output(
            prep_path, xlsx_path, planning_path=planning_path,
            name_aliases=name_aliases_map, recalc=True,
        )
        errors.extend(xlsx_report["errors"])
        warnings.extend(xlsx_report["warnings"])

        pdf_b64 = None
        if not podium_error:
            pdf_report = gen_lib.verify_podium_pdf(pdf_path)
            errors.extend(pdf_report["errors"])
            warnings.extend(pdf_report["warnings"])

            if not skip_ai_review:
                vision_report = gen_lib.ai_visual_review(pdf_path)
                # Toujours en warnings (voir docstring) : la relecture IA
                # n'est jamais bloquante a elle seule.
                warnings.extend(f"Relecture IA : {issue}" for issue in vision_report["issues"])
                warnings.extend(vision_report["warnings"])

            with open(pdf_path, "rb") as f:
                pdf_b64 = base64.b64encode(f.read()).decode("ascii")

        with open(xlsx_path, "rb") as f:
            xlsx_b64 = base64.b64encode(f.read()).decode("ascii")

        status = "blocked" if errors else "ok"
        return JSONResponse({
            "status": status,
            "week": week,
            "errors": errors,
            "warnings": warnings,
            "employee_productivity": employee_productivity,
            "xlsx_filename": xlsx_name,
            "xlsx_base64": xlsx_b64,
            "xlsx_mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "pdf_filename": pdf_name if pdf_b64 else None,
            "pdf_base64": pdf_b64,
            "pdf_mime": "application/pdf",
        })
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération vérifiée : {e}")
