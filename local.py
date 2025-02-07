import pandas as pd
import typst

def load_excel_files():
    """Loads Excel files from local storage."""
    df_sp = pd.read_excel("sp2.xlsx", sheet_name=0).fillna("")
    df_dl = pd.read_excel("deelproject.xlsx", sheet_name=0).fillna("")
    df_ac = pd.read_excel("actie.xlsx", sheet_name=0).fillna("")


    return df_sp, df_dl, df_ac


def extract_project_details(df_sp, project_name):
    """Extracts relevant details of the selected strategic project."""
    row_index = df_sp[df_sp["Naam Strategisch project"] == project_name].index[0]
    return {
        "coordinator": df_sp.loc[row_index, "Coördinator"],
        "medewerkers": df_sp.loc[row_index, "Welke actoren zijn initiatiefnemer binnen het strategisch project"],
        "partners": df_sp.loc[row_index, "Welke actoren zijn initiatiefnemer binnen het strategisch project"],
        "start_date": df_sp.loc[row_index, "Startdatum strategisch project"],
        "end_date": df_sp.loc[row_index, "Einddatum strategisch project"],
        "beschrijving": df_sp.loc[row_index, "Beschrijving"],
        "update": df_sp.loc[row_index, "Halfjaarlijkse update voortgang"],
        "grote_foto": df_sp.loc[row_index,"Grote foto\n"],
        "kleine_foto": df_sp.loc[row_index,"Kleine foto\n"]
    }

def filter_subprojects(df_dl, project_name):
    """Filters subprojects related to the selected strategic project."""
    return df_dl[df_dl["Naam strategisch project"].str.contains(project_name, regex=False, na=False, case=False)]

def generate_report_text(project_details, df_dl, df_ac, project_name):
    """Generates the Typst text for the report."""
    text = f"""
    #import "vlaamse_rand.typ": conf
    #show: conf
    
    #image("images/Picture1.jpg") 
    #place(right + top, image("images/Picture2.jpg", width: 40%))
    #text(size: 36pt, weight: 700, [{project_name}])
    
    {project_details['start_date']} - {project_details['end_date']}
    
    #box(image("images/Picture3.png"), height: 15pt) #text(size: 20pt, fill: orange, weight: 700, [{project_details['coordinator']}])
    #place(left + bottom, image("images/Picture4.png"))
    #pagebreak()
    
    = {project_name}
    
    *Medewerkers:* {project_details['medewerkers']}
    *Coördinerende organisatie en partners:* {project_details['partners']}

    // #grid(
    //     columns: (2fr,1fr),
    //     rect([#align(center+horizon,[{project_details['grote_foto']}])], width: 100%, height: 30%, fill: rgb("#D96422")),
    //     rect([#align(center+horizon,[{project_details['kleine_foto']}])], width: 100%, height: 30%, fill: rgb("#6E8C2B"))
    // )

    #grid(
    columns: (2fr, 1fr),
    image("images/try1.jpg", height: 30%),
    image("images/try2.jpg", fit: "cover", height: 30%)
    )

    {project_details['beschrijving']}

    {project_details['update']}

    #pagebreak()
    """

    for _, deelproject in df_dl.iterrows():
        deelproject_naam = deelproject["Naam deelproject"]
        deelproject_categorie = deelproject["Categorie deelproject"]
        deelproject_omschrijving = deelproject["Omschrijving deelproject"]
        update = deelproject["Halfjaarlijkse update voortgang"]
        promon = deelproject["PROMON-nummer"]

        partners = deelproject["Welke actoren zijn actief in de stuurgroep?"].replace(";",", ").rstrip(", ")

        start = deelproject["Startdatum deelproject"].strftime("%Y-%m-%d")
        einde = deelproject["Einddatum deelproject"].strftime("%Y-%m-%d")

        if einde =="":
            einde = "nu"

        grote_foto = "None"
        kleine_foto = "None"
        
        df_ac_temp = df_ac[
            df_ac["Deze actie linken we aan deelproject"].str.contains(
                deelproject_naam, regex=False, na=False, case=False
            )
        ]

        text += f"""
        == {deelproject_naam}
        """
        if start != "":
             text += f"""
        *{start} tot {einde}*
        """
        if promon != "":
             text += f"""
            *Promon-nummer: {promon}*
            """
             
        if partners != "":
             text += f"""
            *Partners: {partners}*
            """
             
        text += f"""
        *Categorie deelproject: {deelproject_categorie}*
        """

        if grote_foto != None and kleine_foto != None:
            text += f"""
            #grid(
                columns: (2fr,1fr),
                rect([#align(center+horizon,[{grote_foto}])], width: 100%, height: 30%, fill: rgb("#D96422")),
                rect([#align(center+horizon,[{kleine_foto}])], width: 100%, height: 30%, fill: rgb("#6E8C2B"))
            )
            """
        elif grote_foto:
            text += f"""
            #grid(
               columns: 1fr,
                rect([#align(center+horizon,[{grote_foto}])], width: 100%, height: 30%, fill: rgb("#D96422")),
            )
            """

        text += f"""
        {deelproject_omschrijving}

        """
        if update:
            text+= f"#rect([{update}], width: 100%, inset: 12pt, fill:rgb(\"#F4CCB6\"), )"

        if not df_ac_temp.empty:
            text += "\n#align(table(columns: (auto,auto,auto,auto,auto),table.header[*Naam actie*][*Categorie*][*Status*][*Donor*][*Subsidie*],"
            for _, actie in df_ac_temp.iterrows():
                naam_actie = actie['Naam van de actie']
                categorie = actie['Categorie actie - deze vraag is moelijker om te doorlopen, maar is voor redenen van programmeren een stuk eenvoudiger.']
                status = actie['Status actie']

                match status:
                    case "Op schema":
                        status = "#circle(radius: 7pt, stroke: none, fill: green)"
                    case "Geblokkeerd":
                        status = "#circle(radius: 7pt, stroke: none, fill: red)"
                    case "Afgerond":
                        status = "#circle(radius: 7pt, stroke: none, fill: blue)"

                donor = actie['Van wie heeft u de subsidie gekregen?\n']
                subsidie = actie['Hoeveel euro in subsidies heeft u gekregen?\n']
                
                text += f"\n[{naam_actie}],[{categorie}],[{status}], [{donor}], [€ {subsidie}],//"

            text += "\n))"
        text+= "\n#pagebreak()"

    return text

def generate_pdf(text, filename):
    """Writes the Typst text to a file and generates a PDF."""
    with open("output.typ", "w", encoding="utf-8") as file:
        file.write(text)
    pdf = typst.compile("output.typ")
    with open(filename, "wb") as f:
        f.write(pdf)
    print(f"PDF generated: {filename}")

def main():
    df_sp, df_dl, df_ac = load_excel_files()
    project_name = "Leve(n)de Woluwe"
    project_details = extract_project_details(df_sp, project_name)
    df_dl_filtered = filter_subprojects(df_dl, project_name)
    text = generate_report_text(project_details, df_dl_filtered, df_ac, project_name)
    generate_pdf(text, f"{project_name}.pdf")

if __name__ == "__main__":
    main()
