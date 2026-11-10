from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


OUTPUT = "/home/kong-bunyim/Documents/workshop/demo/Spring_Boot_MySQL_Setup_Guide.docx"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color="D9D9D9", size="6"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:color"), color)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)


def set_keep_with_next(paragraph, value=True):
    p_pr = paragraph._p.get_or_add_pPr()
    node = p_pr.find(qn("w:keepNext"))
    if value and node is None:
        node = OxmlElement("w:keepNext")
        p_pr.append(node)
    elif not value and node is not None:
        p_pr.remove(node)


def set_keep_lines(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    if p_pr.find(qn("w:keepLines")) is None:
        p_pr.append(OxmlElement("w:keepLines"))


def add_code_block(doc, code, font_size=8.5):
    p = doc.add_paragraph()
    p.style = doc.styles["Code Block"]
    run = p.add_run(code)
    run.font.name = "Consolas"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
    run.font.size = Pt(font_size)
    set_keep_lines(p)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)
    return p


doc = Document()
section = doc.sections[0]
section.page_width = Inches(8.5)
section.page_height = Inches(11)
section.top_margin = Inches(0.68)
section.bottom_margin = Inches(0.68)
section.left_margin = Inches(0.78)
section.right_margin = Inches(0.78)
section.header_distance = Inches(0.3)
section.footer_distance = Inches(0.35)
doc.settings.odd_and_even_pages_header_footer = True

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Aptos"
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor(35, 42, 52)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.08

title_style = styles["Title"]
title_style.font.name = "Aptos Display"
title_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Display")
title_style.font.size = Pt(25)
title_style.font.bold = True
title_style.font.color.rgb = RGBColor(0, 0, 0)
title_style.paragraph_format.space_after = Pt(5)
title_p_pr = title_style.element.get_or_add_pPr()
title_border = title_p_pr.find(qn("w:pBdr"))
if title_border is not None:
    title_p_pr.remove(title_border)

subtitle = styles["Subtitle"]
subtitle.font.name = "Aptos"
subtitle.font.size = Pt(11.5)
subtitle.font.color.rgb = RGBColor(75, 85, 99)
subtitle.paragraph_format.space_after = Pt(14)

for style_name, size, before, after in (
    ("Heading 1", 16, 14, 6),
    ("Heading 2", 12.5, 10, 4),
):
    style = styles[style_name]
    style.font.name = "Aptos Display"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Display")
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.paragraph_format.space_before = Pt(before)
    style.paragraph_format.space_after = Pt(after)
    style.paragraph_format.keep_with_next = True

for list_style_name in ("List Bullet", "List Number"):
    style = styles[list_style_name]
    style.font.name = "Aptos"
    style.font.size = Pt(10.5)
    style.paragraph_format.space_after = Pt(3)
    style.paragraph_format.left_indent = Inches(0.25)
    style.paragraph_format.first_line_indent = Inches(-0.15)

code_style = styles.add_style("Code Block", 1)
code_style.font.name = "Consolas"
code_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
code_style.font.size = Pt(8.5)
code_style.font.color.rgb = RGBColor(30, 41, 59)
code_style.paragraph_format.left_indent = Inches(0.16)
code_style.paragraph_format.right_indent = Inches(0.16)
code_style.paragraph_format.space_before = Pt(3)
code_style.paragraph_format.space_after = Pt(8)
code_style.paragraph_format.line_spacing = 1.0
p_pr = code_style.element.get_or_add_pPr()
shd = OxmlElement("w:shd")
shd.set(qn("w:fill"), "F3F6F9")
p_pr.append(shd)
spacing = p_pr.find(qn("w:spacing"))
if spacing is None:
    spacing = OxmlElement("w:spacing")
    p_pr.append(spacing)
spacing.set(qn("w:before"), "80")
spacing.set(qn("w:after"), "120")

header = section.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
header_run = header.add_run("DEMO PROJECT GUIDE")
header_run.font.name = "Aptos"
header_run.font.size = Pt(8)
header_run.font.bold = True
header_run.font.color.rgb = RGBColor(75, 85, 99)

footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
footer_run = footer.add_run("Spring Boot MySQL Setup")
footer_run.font.name = "Aptos"
footer_run.font.size = Pt(8)
footer_run.font.color.rgb = RGBColor(100, 110, 120)

even_header = section.even_page_header.paragraphs[0]
even_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
even_header_run = even_header.add_run("DEMO PROJECT GUIDE")
even_header_run.font.name = "Aptos"
even_header_run.font.size = Pt(8)
even_header_run.font.bold = True
even_header_run.font.color.rgb = RGBColor(75, 85, 99)

even_footer = section.even_page_footer.paragraphs[0]
even_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
even_footer_run = even_footer.add_run("Spring Boot MySQL Setup")
even_footer_run.font.name = "Aptos"
even_footer_run.font.size = Pt(8)
even_footer_run.font.color.rgb = RGBColor(100, 110, 120)

doc.add_paragraph("Spring Boot MySQL Setup Guide", style="Title")
doc.add_paragraph("Demo project quick reference", style="Subtitle")
intro = doc.add_paragraph()
intro.add_run(
    "This guide lists the dependencies and configuration required to connect the demo "
    "Spring Boot project to MySQL and let Hibernate create the category table. Copy the "
    "examples, replace the credential placeholders, and restart the application."
)

doc.add_heading("Dependency checklist", level=1)
table = doc.add_table(rows=1, cols=3)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
widths = (Inches(2.35), Inches(3.55), Inches(1.05))
headers = ("Dependency", "Purpose", "Use")
for idx, cell in enumerate(table.rows[0].cells):
    cell.width = widths[idx]
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    set_cell_shading(cell, "1F4E78")
    set_cell_margins(cell, top=120, bottom=120)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if idx < 2 else WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(headers[idx])
    r.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)
    r.font.size = Pt(9.5)
set_repeat_table_header(table.rows[0])

dependency_rows = (
    ("spring-boot-starter-webmvc", "Build MVC controllers and REST endpoints.", "Required"),
    ("spring-boot-starter-data-jpa", "Adds Jakarta Persistence, Hibernate, and Spring Data repositories.", "Required"),
    ("mysql-connector-j", "Provides the JDBC driver used to connect Java to MySQL.", "Runtime"),
    ("spring-boot-devtools", "Restarts the application automatically while developing.", "Optional"),
    ("spring-boot-starter-test", "Provides JUnit and Spring testing support.", "Test only"),
    ("spring-boot-starter", "Provides core Spring Boot features. Other starters already include it.", "Remove duplicate"),
)
for row_index, values in enumerate(dependency_rows, start=1):
    cells = table.add_row().cells
    for idx, value in enumerate(values):
        cells[idx].width = widths[idx]
        cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(cells[idx], top=100, bottom=100)
        if row_index % 2 == 0:
            set_cell_shading(cells[idx], "F5F8FB")
        p = cells[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if idx < 2 else WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(value)
        r.font.size = Pt(9)
        if idx == 0:
            r.font.name = "Consolas"
            r._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
set_table_borders(table)

pom_heading = doc.add_heading("Dependencies for pom xml", level=1)
pom_heading.paragraph_format.page_break_before = True
doc.add_paragraph(
    "Place this block inside the project element in pom.xml. Spring Boot manages the dependency "
    "versions, so these entries do not need version numbers."
)
pom_code = """<dependencies>
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-webmvc</artifactId>
    </dependency>

    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-data-jpa</artifactId>
    </dependency>

    <dependency>
        <groupId>com.mysql</groupId>
        <artifactId>mysql-connector-j</artifactId>
        <scope>runtime</scope>
    </dependency>

    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-devtools</artifactId>
        <optional>true</optional>
    </dependency>

    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-test</artifactId>
        <scope>test</scope>
    </dependency>
</dependencies>"""
add_code_block(doc, pom_code, font_size=8)

database_heading = doc.add_heading("Create the MySQL database", level=1)
database_heading.paragraph_format.page_break_before = True
doc.add_paragraph(
    "Sign in to MySQL as an administrator, then create a database and a dedicated application user. "
    "Replace the sample password before running the commands."
)
sql_code = """CREATE DATABASE IF NOT EXISTS demo_db
    CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

CREATE USER IF NOT EXISTS 'demo_user'@'localhost'
    IDENTIFIED BY 'replace_with_a_strong_password';

GRANT ALL PRIVILEGES ON demo_db.*
    TO 'demo_user'@'localhost';

FLUSH PRIVILEGES;"""
add_code_block(doc, sql_code, font_size=8.5)

doc.add_heading("Datasource configuration", level=1)
doc.add_paragraph(
    "Save the following configuration in src/main/resources/application.yaml. Environment variables "
    "keep the database password out of the source code."
)
yaml_code = """spring:
  application:
    name: demo

  datasource:
    url: ${DB_URL:jdbc:mysql://localhost:3306/demo_db}
    username: ${DB_USERNAME:demo_user}
    password: ${DB_PASSWORD}

  jpa:
    hibernate:
      ddl-auto: update
    show-sql: true
    open-in-view: false
    properties:
      hibernate:
        format_sql: true"""
add_code_block(doc, yaml_code, font_size=8.5)

doc.add_paragraph(
    "Spring Boot detects the MySQL driver from the JDBC URL. A driver class name and Hibernate "
    "dialect are normally unnecessary. The update setting creates missing tables and updates the "
    "development schema without deleting existing data."
)

doc.add_heading("Set credentials and run", level=1)
doc.add_paragraph("On Linux or macOS, set the variables in the same terminal that starts the application.")
shell_code = """export DB_USERNAME=demo_user
export DB_PASSWORD='replace_with_a_strong_password'
./mvnw spring-boot:run"""
add_code_block(doc, shell_code, font_size=8.5)

entity_heading = doc.add_heading("Category entity", level=1)
entity_heading.paragraph_format.page_break_before = True
entity_heading.paragraph_format.space_before = Pt(40)
doc.add_paragraph(
    "Hibernate creates the category table only when the class is a valid JPA entity. Use Jakarta's "
    "Id annotation with a lowercase d."
)
entity_code = """package com.example.demo.model;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "category")
public class Category {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String name;
    private String slug;

    public Long getId() { return id; }
    public String getName() { return name; }
    public String getSlug() { return slug; }

    public void setName(String name) { this.name = name; }
    public void setSlug(String slug) { this.slug = slug; }
}"""
add_code_block(doc, entity_code, font_size=8.2)

doc.add_heading("Verify the connection and table", level=1)
doc.add_paragraph(
    "Start the application without compilation or connection errors. Then sign in to MySQL and run:"
)
verify_code = """USE demo_db;
SHOW TABLES;
DESCRIBE category;"""
add_code_block(doc, verify_code, font_size=8.5)

doc.add_heading("Checks when the table is missing", level=1)
for text in (
    "Confirm MySQL is running on localhost port 3306.",
    "Confirm the database name, username, and password match the MySQL account.",
    "Confirm Category has @Entity and jakarta.persistence.Id.",
    "Confirm the application compiles and reaches the Started DemoApplication message.",
    "Confirm spring.jpa.hibernate.ddl-auto is set to update for development.",
):
    add_bullet(doc, text)

warning = doc.add_paragraph()
lead = warning.add_run("Production note  ")
lead.bold = True
warning.add_run(
    "Use database migrations such as Flyway or Liquibase instead of ddl-auto update in production. "
    "Do not commit real database passwords to the repository."
)

doc.core_properties.title = "Spring Boot MySQL Setup Guide"
doc.core_properties.subject = "Dependency, datasource, database, and JPA entity setup for the demo project"
doc.core_properties.author = "Demo Project"
doc.core_properties.keywords = "Spring Boot, MySQL, JPA, Hibernate, Maven"
doc.save(OUTPUT)
print(OUTPUT)
