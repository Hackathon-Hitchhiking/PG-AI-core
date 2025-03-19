SHEMAS_STRUCTURE = """Below is a prompt you can show to users (or LLMs) describing each of the Pydantic dataclasses and how they are structured. It clarifies what each class represents and which fields they contain. You can place this text in documentation or as an in-line comment for easy reference:

⸻

Prompt: Dataclass Structure Explanation

Below is an overview of the custom data models used for managing and describing slides, shapes, and presentation metadata in a PowerPoint-like context. Each class is a Pydantic dataclass, enforcing type validation and containing detailed descriptions of fields:

⸻

FillType

An enumeration (IntEnum) listing the possible fill types for shapes and backgrounds:
	•	SOLID (1): Solid fill
	•	THEME (2): Theme-based color fill
	•	PICTURE (3): Picture-based fill
	•	PATTERN (4): Pattern-based fill

⸻

Coordinates

Represents the position and size of an element, in inches.
	•	x (Optional[float], default None): Horizontal position
	•	y (Optional[float], default None): Vertical position
	•	width (float, required): Width of the element
	•	height (float, required): Height of the element

⸻

FontStyle

Describes text styling attributes.
	•	name (str, default "Calibri"): The font name
	•	size (float, default 12.0): Font size in points
	•	color (str, default "#000000"): Text color in HEX format (e.g., #FFFFFF)
	•	bold (bool, default False): Whether text is bold
	•	italic (bool, default False): Whether text is italic
	•	underline (bool, default False): Whether text is underlined
	•	alignment (Literal "left", "center", "right", "justify", default "left"): Text alignment

⸻

Background

Represents the background of a slide.
	•	type (Literal "color", "image", "pattern", "none", default "none"): Background type
	•	value (Optional[str]): Either a HEX color code (e.g., #FF00FF), image path, or pattern name, depending on type
	•	transparency (float, default 1.0): Transparency level (0.0 is fully opaque, 1.0 is fully transparent)

⸻

SlideLayout

Details about a single slide layout in the presentation.
	•	name (str): Machine-readable layout name
	•	display_name (str): Human-friendly layout name
	•	idx (int): Zero-based index of the layout in the presentation

⸻

TextContent

Encapsulates the textual data inside a text shape.
	•	text (str, required): Actual text content
	•	bullet (bool, default False): Whether to display bullets
	•	bullet_level (int, default 0): Nesting level for bullets (0 = no indent, 1 = first indent, etc.)

⸻

TextShape

Represents a text box shape on a slide.
	•	type (Literal "text", default = "text"): Distinguishes this shape as text
	•	id (str, required): Unique shape identifier
	•	content (TextContent): Text data (content, bulleting)
	•	coordinates (Coordinates): Position and size
	•	style (FontStyle): Font styling (bold, italics, etc.)

⸻

ImageShape

Represents an image placed on a slide.
	•	type (Literal "image", default = "image"): Distinguishes this shape as an image
	•	id (str, required): Unique shape identifier
	•	path (str, required): Path to the image file (or a base64 string)
	•	coordinates (Coordinates): Position and size
	•	crop (Optional[Coordinates], default None): Crop area, if any
	•	brightness (float, default 1.0): Multiplier for image brightness (0.0–2.0)
	•	contrast (float, default 1.0): Multiplier for image contrast (0.0–2.0)

⸻

ChartSeries

Represents a data series within a chart.
	•	name (str): Name of the series (e.g., “Q1 Sales”)
	•	values (list of float): Numeric data points

⸻

ChartData

Holds the categories and series for a chart.
	•	categories (list of str): Label set for the X-axis (or equivalent)
	•	series (list of ChartSeries): Series collection containing name/value data

⸻

ChartShape

Represents a chart on a slide.
	•	type (Literal "chart", default = "chart"): Distinguishes this shape as a chart
	•	id (str, required): Unique shape identifier
	•	chart_type (str): E.g., "bar", "line", "pie", "area"
	•	data (ChartData): Chart data (categories, series)
	•	coordinates (Coordinates): Chart position and size
	•	title (Optional[str], default None): Optional chart title

⸻

TableCell

Details about a single cell in a table.
	•	row (int): Row index
	•	col (int): Column index
	•	content (str): Cell text content
	•	span_rows (int, default 1): Vertical cell span (merge)
	•	span_cols (int, default 1): Horizontal cell span (merge)

⸻

TableShape

Represents a table shape on a slide.
	•	type (Literal "table", default = "table"): Distinguishes this shape as a table
	•	id (str, required): Unique shape identifier
	•	rows (int): Number of rows
	•	cols (int): Number of columns
	•	cells (list of TableCell): Collection of all cells in the table
	•	coordinates (Coordinates): Table position and size

⸻

GeometricShape

Represents a simple auto-shape (rectangle, ellipse, arrow, etc.).
	•	type (Literal "shape", default = "shape"): Distinguishes this shape as geometric
	•	id (str, required): Unique shape identifier
	•	shape_type (str, default = "rectangle"): Type name of the shape (e.g., "rounded_rectangle", "ellipse")
	•	coordinates (Coordinates): Position and size
	•	fill (Optional[str], default None): HEX color or None if no fill
	•	outline (Optional[str], default None): HEX color or None if no outline
	•	outline_width (float, default 0.0): Outline thickness in inches

⸻

Shape

A union type that can be one of the following shapes:
	•	TextShape
	•	ImageShape
	•	ChartShape
	•	TableShape
	•	GeometricShape

Pydantic will attempt to validate an incoming shape object against one of these five classes.

⸻

SlideDesign

Theme-level design info from a slide.
	•	theme_colors (list of str): Theme’s color scheme
	•	theme_fonts (dict[str, str]): Mapping of font roles (latin, complex, etc.) to typeface
	•	default_margin (float, default = 0.5): A default margin in inches

⸻

Slide

Describes one slide in the presentation.
	•	number (int): 1-based index of the slide
	•	layout (SlideLayout): Slide layout details
	•	background (Background): Slide background settings
	•	shapes (list of Shape): All recognized shapes (text, image, chart, table, geometry)
	•	design (SlideDesign): Theme info extracted from the slide

⸻

PresentationMeta

Core properties of a presentation (metadata).
	•	title (Optional[str]): Document title
	•	author (Optional[str]): Document author
	•	created (Optional[datetime]): Creation timestamp
	•	modified (Optional[datetime]): Last modified timestamp
	•	template (Optional[str]): If a template was used

⸻

PresentationInfo

High-level overview of a presentation, including:
	•	meta (PresentationMeta): Core presentation metadata
	•	dimensions (Coordinates): Slide width/height (inches)
	•	slides_count (int): Number of slides
	•	slides (list of Slide): Detailed info about each slide
	•	master_layouts (list of SlideLayout): All layout definitions available

⸻

Use these definitions to fully understand the model structure and how the data is validated and stored. For example, a text shape must include type = "text", an ID, a TextContent object for the text, Coordinates for placement, and a FontStyle for text appearance. Meanwhile, a chart shape requires chart data (ChartData) with categories and series, along with a recognized chart type.

You can share these definitions with any calling application or function-calling LLM to ensure consistent usage and better error handling.
"""