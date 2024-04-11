# Load necessary libraries
library(officer)
library(magrittr)

# Create a new Word document
doc <- read_docx()

# Add a logo
doc <- doc %>%
  body_add_img(src = "escudo_basico/PNG/Escudo Color TL.png", width = 2.8, height = 1, pos = "after") %>%
  body_add_par("", style = "Normal")

# Add a recommendation letter
fp_normal <- fp_text(font.family = "Arial", font.size = 10)
fp_italic <- update(fp_normal, italic = TRUE)
fp_bold <- update(fp_normal, bold = TRUE)
fp_link <- update(fp_normal, underlined = TRUE, color = "blue")

doc <- doc %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(fpar(ftext("Burgos, 11 de abril de 2024.", fp_normal), fp_p = fp_par(text.align = "right"))) %>% 
  body_add_fpar(fpar(ftext("", fp_text(font.family = "Arial", font.size = 11, bold = TRUE)), fp_p = fp_par(text.align = "right"))) %>% 
  #body_add_par(fpar(ftext("A quien corresponda,", fp_normal)), style = "Normal") %>%
  body_add_fpar(
    block_list(
      fpar(ftext("A quien corresponda,", fp_normal)
      )),
    style = "Normal"
  ) %>% 
  body_add_par("", style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(fpar(ftext("Me dirijo a usted para recomendarle a [NOMBRE RECOMENDADO], que ha solicitado la admisión en vuestra institución.", fp_normal), fp_p = fp_par(line_spacing = 1.25))) %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(fpar(ftext("Como director del Área de Lenguajes y Sistemas Informáticos del Departamento de Ingeniería Informática de la Universidad de Burgos, conozco a la candidata [NOMBRE RECOMENDADO] por sus estudios cursados en el Grado de la Ingeniería de la Salud de nuestra institución. En las dos asignaturas en que he tenido a la candidata de alumna, su dedicación y constancia hacia los estudios han sido más que notables. En particular, destacan sus habilidades en el tratamiento de datos y en el manejo de herramientas informáticas y de computación.", fp_normal), fp_p = fp_par(line_spacing = 1.25))) %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(
    block_list(
      fpar(ftext("Además de haber sido su profesor, he tenido el placer de ser su tutor de trabajo de fin de grado, titulado: ", fp_normal), fp_p = fp_par(line_spacing = 1.25), 
      ftext('"Respuesta Sensorial Meridiana Autónoma (ASMR): Análisis de la Experiencia Global y Desarrollo de una Aplicación de Recomendación supervisada de ASMR"', fp_text(font.family = "Arial", font.size = 10, italic = TRUE, underlined = TRUE)), 
      ftext("; a ser presentado en la convocatoria de junio del curso 2023-24. El trabajo posee una calidad técnica y científica más que destacable, permitiéndole demostrar sus capacidades como ingeniera de la salud. El trabajo está públicamente accesible en el repositorio Github:", fp_normal),
      ftext("https://github.com/rocioalonsolh/tfg-asmr", fp_link)
    )),
    style = "Normal"
  ) %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(fpar(ftext("Su interés en la Minería de Datos, la Inteligencia Artificial y la Ciberseguridad ha sido una constante a lo largo de todos sus estudios universitarios, culminando con el previamente mencionado trabajo de fin de grado donde se familiarizó con las técnicas de programación en lenguajes como R y Python, preprocesamiento y análisis de datos y diseño e implementación de aplicaciones web tipo Shiny.", fp_normal), fp_p = fp_par(line_spacing = 1.25))) %>%
  body_add_fpar(fpar(ftext("", fp_normal), fp_p = fp_par(line_spacing = 1.25))) %>%
  body_add_fpar(fpar(ftext("Por todos estos motivos, recomiendo a [NOMBRE RECOMENDADO] como candidata por su destacado interés en la materia, alta formación, responsabilidad y eficiencia en el ámbito académico.", fp_normal), fp_p = fp_par(line_spacing = 1.25))) %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(
    block_list(
      fpar(ftext("", fp_normal)
      )),
    style = "Normal"
  ) %>% 
  #body_add_par("Atentamente,", style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  body_add_fpar(fpar(ftext("Atentamente, ", fp_normal), fp_p = fp_par(text.align = "left"))) %>% 
  body_add_par("", style = "Normal") %>%
  body_add_par("", style = "Normal") %>%
  
  body_add_fpar(fpar(ftext("Dr. Antonio Canepa Oneto", fp_text(font.family = "Arial", font.size = 11, bold = TRUE)), fp_p = fp_par(text.align = "left"))) %>% 
  body_add_fpar(fpar(ftext("Director Área de Lenguajes y Sistemas Informáticos", fp_italic), fp_p = fp_par(text.align = "left"))) %>% 
  body_add_fpar(fpar(ftext("Escuela Politécnica Superior", fp_italic), fp_p = fp_par(text.align = "left"))) %>% 
  body_add_fpar(fpar(ftext("Universidad de Burgos", fp_italic), fp_p = fp_par(text.align = "left")))

# Save the Word document
print(doc, target = "Recommendation_Letter_[RECOMENDADO].docx")
