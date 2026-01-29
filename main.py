"""
Beam Report Generator Using PyLaTeX
Generates a professional PDF report for simply supported beam analysis
Uses pre-calculated SFD and BMD data from Excel
"""

import pandas as pd
import numpy as np
from pylatex import Document, Section, Subsection, Figure, NoEscape, Package
from pylatex.utils import bold
import os


def read_beam_data(excel_path):
    """
    Read pre-calculated beam analysis data from Excel file
    
    Args:
        excel_path: Path to the Excel file containing x, Shear force, and Bending Moment data
        
    Returns:
        DataFrame with beam analysis data
    """
    try:
        df = pd.read_excel(excel_path)
        print(f"Successfully read {len(df)} data points from Excel file")
        print(f"Columns: {df.columns.tolist()}")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def create_tikz_sfd(x_data, y_data):
    """
    Create TikZ code for Shear Force Diagram
    
    Args:
        x_data: List of x coordinates
        y_data: List of shear force values
        
    Returns:
        LaTeX string with TikZ code
    """
    coordinates = " ".join([f"({x},{y})" for x, y in zip(x_data, y_data)])
    
    tikz_code = r"""
\begin{center}
\begin{tikzpicture}
    \begin{axis}[
        width=12cm,
        height=7cm,
        xlabel={Distance along beam (m)},
        ylabel={Shear Force (kN)},
        grid=major,
        legend pos=north east,
        axis lines=middle,
    ]
    \addplot[blue, thick, smooth] coordinates {
        """ + coordinates + r"""
    };
    \addlegendentry{Shear Force}
    \end{axis}
\end{tikzpicture}
\end{center}
"""
    return tikz_code


def create_tikz_bmd(x_data, y_data):
    """
    Create TikZ code for Bending Moment Diagram
    
    Args:
        x_data: List of x coordinates
        y_data: List of bending moment values
        
    Returns:
        LaTeX string with TikZ code
    """
    coordinates = " ".join([f"({x},{y:.2f})" for x, y in zip(x_data, y_data)])
    
    tikz_code = r"""
\begin{center}
\begin{tikzpicture}
    \begin{axis}[
        width=12cm,
        height=7cm,
        xlabel={Distance along beam (m)},
        ylabel={Bending Moment (kN·m)},
        grid=major,
        legend pos=north east,
        axis lines=middle,
    ]
    \addplot[red, thick, smooth] coordinates {
        """ + coordinates + r"""
    };
    \addlegendentry{Bending Moment}
    \end{axis}
\end{tikzpicture}
\end{center}
"""
    return tikz_code


def create_data_table(beam_df):
    """
    Create LaTeX tabular code for beam data table
    
    Args:
        beam_df: DataFrame containing beam analysis data
        
    Returns:
        LaTeX string with tabular code
    """
    table_code = r"""
\begin{center}
\begin{tabular}{|c|c|c|}
\hline
\textbf{Position (m)} & \textbf{Shear Force (kN)} & \textbf{Bending Moment (kN·m)} \\
\hline
"""
    
    step = 1 if len(beam_df) <= 15 else 2
    
    for idx, row in beam_df.iloc[::step].iterrows():
        table_code += f"{row['x']:.1f} & {row['Shear force']:.2f} & {row['Bending Moment']:.2f} \\\\\n\\hline\n"
    
    table_code += r"""
\end{tabular}
\end{center}
"""
    return table_code


def generate_report(beam_df, beam_image_path, output_path):
    """
    Generate the complete PDF report
    
    Args:
        beam_df: DataFrame containing beam analysis data
        beam_image_path: Path to beam diagram image
        output_path: Path where PDF should be saved
    """
    geometry_options = {"margin": "1in"}
    doc = Document(geometry_options=geometry_options)
    
    doc.packages.append(Package('graphicx'))
    doc.packages.append(Package('pgfplots'))
    doc.packages.append(Package('tikz'))
    doc.preamble.append(NoEscape(r'\pgfplotsset{compat=1.18}'))
    
    doc.preamble.append(NoEscape(r'\title{Beam Analysis Report \\ Simply Supported Beam}'))
    doc.preamble.append(NoEscape(r'\author{Generated using PyLaTeX}'))
    doc.preamble.append(NoEscape(r'\date{\today}'))
    doc.append(NoEscape(r'\maketitle'))
    
    doc.append(NoEscape(r'\tableofcontents'))
    doc.append(NoEscape(r'\newpage'))
    
    with doc.create(Section('Introduction')):
        doc.append('This report presents the structural analysis of a simply supported beam. ')
        doc.append('The beam consists of a pinned support on the left end and a roller support on the right end, ')
        doc.append('allowing for horizontal movement while providing vertical support.')
        doc.append(NoEscape(r'\par\vspace{0.5cm}'))
        
        doc.append('The analysis includes the complete shear force distribution and bending moment distribution ')
        doc.append('along the length of the beam.')
        doc.append(NoEscape(r'\par\vspace{0.5cm}'))
        
        doc.append(bold('Data Source: '))
        doc.append('Analysis data has been imported from the Excel file: Force.xlsx')
        doc.append(NoEscape(r'\par\vspace{0.5cm}'))
        
        if os.path.exists(beam_image_path):
            with doc.create(Figure(position='h!')) as fig:
                fig.add_image(beam_image_path, width='400px')
                fig.add_caption('Simply Supported Beam Configuration')
        
        doc.append(NoEscape(r'\par\vspace{0.5cm}'))
        
        beam_length = beam_df['x'].max() - beam_df['x'].min()
        doc.append(bold('Beam Properties:'))
        doc.append(NoEscape(r'\par'))
        doc.append(f'Total Beam Length: {beam_length:.1f} m')
        doc.append(NoEscape(r'\par'))
        doc.append(f'Number of Analysis Points: {len(beam_df)}')
    
    doc.append(NoEscape(r'\newpage'))
    
    with doc.create(Section('Analysis Data')):
        doc.append('The following table shows the calculated shear force and bending moment values ')
        doc.append('at various positions along the beam:')
        doc.append(NoEscape(r'\par\vspace{0.5cm}'))
        
        doc.append(NoEscape(create_data_table(beam_df)))
        
        doc.append(NoEscape(r'\par\vspace{0.5cm}'))
        
        max_shear = beam_df['Shear force'].abs().max()
        max_moment = beam_df['Bending Moment'].max()
        max_moment_location = beam_df.loc[beam_df['Bending Moment'].idxmax(), 'x']
        
        doc.append(bold('Key Analysis Results:'))
        doc.append(NoEscape(r'\par'))
        doc.append(f'Maximum Shear Force: {max_shear:.2f} kN')
        doc.append(NoEscape(r'\par'))
        doc.append(f'Maximum Bending Moment: {max_moment:.2f} kN·m')
        doc.append(NoEscape(r'\par'))
        doc.append(f'Location of Maximum Moment: {max_moment_location:.2f} m from left support')
    
    doc.append(NoEscape(r'\newpage'))
    
    with doc.create(Section('Structural Analysis Diagrams')):
        
        with doc.create(Subsection('Shear Force Diagram (SFD)')):
            doc.append('The shear force diagram shows the variation of shear force along the length of the beam. ')
            doc.append('Shear force represents the internal force acting perpendicular to the beam axis.')
            doc.append(NoEscape(r'\par\vspace{0.5cm}'))
            
            x_data = beam_df['x'].tolist()
            shear_data = beam_df['Shear force'].tolist()
            
            doc.append(NoEscape(create_tikz_sfd(x_data, shear_data)))
            
            doc.append(NoEscape(r'\par\vspace{0.3cm}'))
            doc.append(f'The shear force varies from {min(shear_data):.2f} kN to {max(shear_data):.2f} kN along the beam.')
        
        doc.append(NoEscape(r'\vspace{1.5cm}'))
        
        with doc.create(Subsection('Bending Moment Diagram (BMD)')):
            doc.append('The bending moment diagram shows the variation of bending moment along the length of the beam. ')
            doc.append('Bending moment represents the internal moment that causes the beam to bend.')
            doc.append(NoEscape(r'\par\vspace{0.5cm}'))
            
            moment_data = beam_df['Bending Moment'].tolist()
            
            doc.append(NoEscape(create_tikz_bmd(x_data, moment_data)))
            
            doc.append(NoEscape(r'\par\vspace{0.3cm}'))
            doc.append(f'The maximum bending moment is {max(moment_data):.2f} kN·m, ')
            doc.append(f'occurring at {max_moment_location:.2f} m from the left support.')
    
    doc.append(NoEscape(r'\newpage'))
    with doc.create(Section('Conclusion')):
        doc.append('The structural analysis of the simply supported beam has been completed. ')
        doc.append('The shear force and bending moment diagrams provide essential information for:')
        doc.append(NoEscape(r'\par\vspace{0.3cm}'))
        doc.append(NoEscape(r'\begin{itemize}'))
        doc.append(NoEscape(r'\item Determining critical sections for design'))
        doc.append(NoEscape(r'\item Calculating required beam dimensions'))
        doc.append(NoEscape(r'\item Selecting appropriate materials'))
        doc.append(NoEscape(r'\item Ensuring structural safety and stability'))
        doc.append(NoEscape(r'\end{itemize}'))
        doc.append(NoEscape(r'\par\vspace{0.3cm}'))
        doc.append('These results form the foundation for detailed structural design and verification.')
    
    try:
        doc.generate_pdf(output_path.replace('.pdf', ''), clean_tex=False)
        print(f"\nReport successfully generated: {output_path}")
    except Exception as e:
        print(f"\nError generating PDF: {e}")
        print("Make sure you have LaTeX installed (MiKTeX, TeX Live, or MacTeX)")


def main():
    """
    Main function to orchestrate report generation
    """
    excel_path = 'data/Force.xlsx'
    beam_image_path = 'images/Beam.png'
    output_path = 'output/report.pdf'
    
    os.makedirs('output', exist_ok=True)
    
    print("Reading beam analysis data from Excel...")
    beam_df = read_beam_data(excel_path)
    
    if beam_df is not None:
        print("\nGenerating report...")
        generate_report(beam_df, beam_image_path, output_path)
        print("\n✓ Report generation complete!")
        print(f"✓ Output saved to: {output_path}")
    else:
        print("\n✗ Failed to read beam data. Report generation aborted.")


if __name__ == "__main__":
    main()