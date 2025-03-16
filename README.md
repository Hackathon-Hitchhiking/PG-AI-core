# PG-AI-Core

PG-AI-Core is a comprehensive AI core framework designed to facilitate the development of text and image generation agents. This project provides a robust infrastructure for integrating various AI models and APIs, enabling seamless generation of content for different applications.

## Features

- **Text Generation**: Supports multiple backends including Hugging Face, OpenAI, LangChain, LlamaCpp, and more.
- **Image Generation**: Integrates with Diffusers, OpenAI DALL-E, Yandex Cloud, and custom APIs.
- **Presentation Generation**: Combines text and image generation capabilities to create comprehensive presentations.
- **Caching and Fallbacks**: Implements caching mechanisms and fallback models to ensure reliability and efficiency.
- **Extensible Architecture**: Easily extendable to support new models and APIs.

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/PG-AI-Core.git
    cd PG-AI-Core
    ```

2. Install dependencies using PDM:
    ```sh
    pdm install
    ```

## Usage

### Text Generation

To generate text using the provided agents, you can use the `TextAgent` class. Below is an example of how to initialize and use the agent:

```python
from agents.text.agent import TextAgent, TextModelConfig

config = TextModelConfig(
    backend="hf",
    model_name="gpt-3.5-turbo",
    device="cuda",
    torch_dtype="float16"
)

agent = TextAgent(config)
title = agent.generate_title("Future of AI in healthcare")
content = agent.generate_content("Technical explanation of neural networks")
print(f"Generated Title: {title}")
print("Generated Content:")
print(content)
```

### Image Generation

To generate images, use the `ImageAgent` class. Here is an example:

```python
from agents.image.schemas import DiffusersConfig, GenerationRequest
from agents.image.agent import ImageAgent

config = DiffusersConfig(
    model_name="stabilityai/stable-diffusion-xl-base-1.0",
    device="cuda",
    torch_dtype="float16"
)

agent = ImageAgent(config)
request = GenerationRequest(
    prompt="A futuristic cityscape at sunset",
    width=1024,
    height=768,
    num_inference_steps=40,
    guidance_scale=7.5,
    seed=42
)

image = agent.generate(request)
image.save("output_image.png")
```

### Presentation Generation

To create a presentation, use the `PresentationAgent` class. Below is an example:

```python
from agents.presentation import PresentationAgent
from agents.text.agent import TextAgent
from agents.image.agent import ImageAgent
from config import CONFIG

text_agent = TextAgent(CONFIG['text_model'])
image_agent = ImageAgent(CONFIG['image_model'])

presentation_agent = PresentationAgent(text_agent, image_agent, CONFIG)
presentation_agent.create_presentation("AI in Healthcare", document_path="input.docx", template_path="template.pptx")
```

## Configuration

The configuration for the models and APIs is managed through the `config.py` file. Update the configuration as needed to match your setup.

```python
# config.py
CONFIG = {
    'text_model': {
        'type': 'api',
        'api_endpoint': 'https://api.example.com/v1/completions',
        'api_key': 'YOUR_API_KEY',
        'model_name': 'gpt-3.5-turbo',
        'local_path': 'models/gpt-3.5-turbo',
        'temperature': 0.7
    },
    'image_model': {
        'type': 'api',
        'api_endpoint': 'https://api.example.com/v1/images/generations',
        'api_key': 'YOUR_API_KEY',
        'model_name': 'dall-e-3',
        'local_path': 'models/dall-e-3',
        'guidance_scale': 7.5
    },
    'templates_dir': 'resources/templates',
    'output_dir': 'output/presentations',
    'prompt_templates': 'resources/prompts.json',
    'default_slide_count': 10,
    'max_retries': 3,
    'timeout': 60
}
```

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

- [Hugging Face](https://huggingface.co/)
- [OpenAI](https://www.openai.com/)
- [Yandex Cloud](https://cloud.yandex.com/)
- [LangChain](https://langchain.com/)
- [Diffusers](https://github.com/huggingface/diffusers)

## Contact

For any inquiries or issues, please contact [Leonid Chesnikov](mailto:leonid.chesnikov@gmail.com).
