# PPT export to PNG
# Hogyan használjuk:
# 1. Minden fájl és mappanevet nevezzünk át ékezet nélkülire! (nagyon fontos)
# 2. telepítsük a comtypes csomagot: pip install comtypes
# 3. A ppt_folder változóban adjuk meg a ppt fájlok mappáját
# 4. Az output_folder változóban adjuk meg a kimeneti mappát
# 5. Csak relatív mappaneveket adjunk meg!
# 6. Futtassuk a kódot
# 7. A PNG-k a kimeneti mappában lesznek

import os
import comtypes.client
import logging
import time

# Set up logging
logging.basicConfig(filename="ppt_export_multiple.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def export_slides_from_ppt(ppt_path, output_folder):
    """Exports slides from a PowerPoint presentation to PNG format into a specified output folder."""
    try:
        print(f"Processing file: {ppt_path}")
        logging.info(f"Processing {ppt_path}")
        print(f"Currently exporting: {os.path.basename(ppt_path)}")

        # Create subfolder in output directory based on the presentation name
        ppt_name = os.path.splitext(os.path.basename(ppt_path))[0]
        ppt_output_folder = os.path.join(output_folder, ppt_name)
        os.makedirs(ppt_output_folder, exist_ok=True)

        # Open PowerPoint application
        print("Attempting to open PowerPoint...")
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # Keep PowerPoint visible
        print("PowerPoint opened successfully.")

        # Add a small delay to ensure PowerPoint is fully started
        time.sleep(2)

        # Open the PowerPoint presentation
        print(f"Opening presentation: {ppt_path}")
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)  # Open in read-only mode
        print(f"Presentation {ppt_path} opened successfully.")

        # Export each slide as PNG
        slide_count = len(presentation.Slides)
        if slide_count == 0:
            print(f"No slides found in presentation: {ppt_path}")
            logging.warning(f"No slides found in {ppt_path}")
        else:
            for i, slide in enumerate(presentation.Slides):
                # Define the slide export path (inside the output subfolder)
                slide_export_path = os.path.join(ppt_output_folder, f"slide_{i+1}.png")
                
                slide.Export(slide_export_path, "PNG")
                logging.info(f"Exported slide {i+1} to {slide_export_path}")
                print(f"Exported slide {i+1} to {slide_export_path}")

                # Add a small delay to ensure the file is saved properly
                time.sleep(0.5)

        # Close the presentation and PowerPoint
        presentation.Close()
        powerpoint.Quit()

        logging.info(f"Successfully exported slides from {ppt_path}")
        print(f"Successfully exported: {os.path.basename(ppt_path)}")

    except Exception as e:
        logging.error(f"Failed to export slides from {ppt_path}: {e}")
        print(f"Error processing {os.path.basename(ppt_path)}: {e}")


def process_presentations(ppt_folder, output_folder):
    """Processes all PowerPoint presentations in a folder and exports slides to the output folder."""
    # Get the current working directory
    working_directory = os.getcwd()

    # Concatenate the working directory with the user-specified folder paths
    ppt_folder_full = os.path.join(working_directory, ppt_folder)
    output_folder_full = os.path.join(working_directory, output_folder)

    print(f"Looking for presentations in: {ppt_folder_full}")
    print(f"Saving PNG files to: {output_folder_full}")

    # Ensure the output folder exists
    if not os.path.exists(output_folder_full):
        os.makedirs(output_folder_full)

    # Loop through all files in the specified folder
    for file_name in os.listdir(ppt_folder_full):
        if file_name.endswith((".ppt", ".pptx")):  # Check for both .ppt and .pptx
            ppt_path = os.path.join(ppt_folder_full, file_name)
            print(f"Found presentation: {file_name}")
            export_slides_from_ppt(ppt_path, output_folder_full)
        else:
            print(f"Skipping non-PPT/PPTX file: {file_name}")


if __name__ == "__main__":
    # User should only provide relative directory paths!
    ppt_folder = "word fajlok es pptk"  # Replace with the relative path to the presentations
    output_folder = "PNG_OUTPUT"  # Replace with the relative path to the output folder

    print(f"Starting export process for presentations in: {ppt_folder}")
    process_presentations(ppt_folder, output_folder)
    print("All presentations have been processed.")