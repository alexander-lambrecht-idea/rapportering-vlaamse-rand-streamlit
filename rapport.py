import streamlit as st
from streamlit_sortables import sort_items
import pandas as pd
import typst


def upload_excel(prompt):
    """Handles file upload and reads the Excel file."""
    file = st.file_uploader(prompt)
    return pd.read_excel(file, sheet_name=0) if file else None


def select_strategic_project(df_sp):
    """Handles the selection of a strategic project."""
    project_list = df_sp["Naam Strategisch project"].tolist()
    return st.radio(
        "Van welk strategisch project wilt u een rapport?", options=project_list
    )


def extract_project_details(df_sp, project_name):
    """Extracts relevant details of the selected strategic project."""
    row_index = df_sp[df_sp["Naam Strategisch project"] == project_name].index[0]
    return {
        "coordinator": df_sp.loc[row_index, "Coördinator"],
        "medewerkers": df_sp.loc[
            row_index,
            "Welke actoren zijn initiatiefnemer binnen het strategisch project",
        ],
        "partners": df_sp.loc[
            row_index,
            "Welke actoren zijn initiatiefnemer binnen het strategisch project",
        ],
        "start_date": df_sp.loc[row_index, "Startdatum strategisch project"],
        "end_date": df_sp.loc[row_index, "Einddatum strategisch project"],
        "beschrijving": df_sp.loc[row_index, "Beschrijving"],
        "update": df_sp.loc[row_index, "Halfjaarlijkse update voortgang"],
        "grote_foto": df_sp.loc[row_index,"Grote foto\n"],
        "kleine_foto": df_sp.loc[row_index,"Kleine foto\n"]
    }


def filter_subprojects(df_dl, project_name):
    """Filters subprojects related to the selected strategic project."""
    return df_dl[
        df_dl["Naam strategisch project"].str.contains(
            project_name, regex=False, na=False, case=False
        )
    ]


def generate_report_text(project_details, df_dl, df_ac, project_name, dl_order):
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
    df_dl["Naam deelproject"] = pd.Categorical(df_dl["Naam deelproject"], categories=dl_order, ordered=True)

# Sort the DataFrame based on the categorical order
    df_dl = df_dl.sort_values("Naam deelproject")
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
    st.download_button("Download uw rapport", pdf, file_name=filename)

def select_order_dl(df_dl):
    st.subheader("Gelieve de deelprojecten te sorteren in de gewilde volgorde")
    deelprojecten = df_dl["Naam deelproject"]
    deelprojecten = sort_items(deelprojecten.to_list())
    return deelprojecten

def main():
    st.title("Welkom bij rapportering van de Vlaamse rand!")
    st.header("Gelieve de excel up te loaden van het niveau strategisch project")
    df_sp = upload_excel(
        "Gelieve de excel up te loaden van het niveau *strategisch project*"
    )

    if df_sp is not None:
        st.header("Gelieve de excel up te loaden van het niveau *deelproject*")
    
        project_name = select_strategic_project(df_sp)
        project_details = extract_project_details(df_sp, project_name)
        df_dl = upload_excel(
            "Gelieve de excel up te loaden van het niveau *deelprojecten*"
        )

        if df_dl is not None:
            dl_order = select_order_dl(df_dl)
            st.header("Gelieve de excel up te loaden van het niveau actie")
    
            df_ac = upload_excel(
                "Gelieve de excel up te loaden van het niveau *acties*"
            )
            if df_ac is not None:
                df_dl_filtered = filter_subprojects(df_dl, project_name)
                text = generate_report_text(
                    project_details, df_dl_filtered, df_ac, project_name, dl_order
                )
                generate_pdf(text, f"{project_name}.pdf")


if __name__ == "__main__":
    main()
