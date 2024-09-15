import os
import comtypes.client
import logging
import time

# Set up logging
logging.basicConfig(filename="ppt_export_simple.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def export_slides_from_ppt(ppt_path):
    """Exports slides from a PowerPoint presentation to PNG format in the same directory as the .ppt/.pptx file."""
    try:
        print(f"Verifying path: {ppt_path}")
        logging.info(f"Processing {ppt_path}")
        print(f"Currently exporting: {os.path.basename(ppt_path)}")

        # Determine the output folder (same directory as ppt/pptx)
        output_folder = os.path.dirname(ppt_path)

        # Open PowerPoint application
        print("Attempting to open PowerPoint...")
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # Make PowerPoint visible (if you want to see the window)
        print("PowerPoint opened successfully.")

        # Add a delay to give PowerPoint time to fully start
        time.sleep(2)

        # Open the PowerPoint presentation in read-only mode
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
                # Define the slide export path (same directory as the .ppt/.pptx file)
                slide_export_path = os.path.join(output_folder, f"slide_{i+1}.png")
                
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


if __name__ == "__main__":
    # Test with the path to the PowerPoint file
    ppt_path = "G:\\jovotajolo\\resources\\1st curse\\word fajlok es pptk\\Horthy1.pptx"  # Replace with your file path

    print(f"Starting export for presentation: {ppt_path}")
    export_slides_from_ppt(ppt_path)
    print("Presentation has been processed.")
