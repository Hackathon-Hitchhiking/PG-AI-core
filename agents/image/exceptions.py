class ImageGenerationError(Exception):
    """Base exception for image generation errors."""

    pass


class ModelInitializationError(ImageGenerationError):
    """Raised when model initialization fails."""

    pass


class ValidationError(ImageGenerationError):
    """Raised when input validation fails."""

    pass


class GenerationError(ImageGenerationError):
    """Raised when image generation fails."""

    pass
