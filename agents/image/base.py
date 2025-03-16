from abc import ABC, abstractmethod

from PIL import Image

from schemas.image import GenerationRequest, ImageModelConfig

class BaseImageModel(ABC):
    """
    Abstract base class for image models.
    """

    @abstractmethod
    def generate(self, request: GenerationRequest) -> Image.Image:
        """
        Generates an image based on a GenerationRequest.

        Parameters
        ----------
        request : GenerationRequest
            Parameters for the image generation.

        Returns
        -------
        PIL.Image.Image
            Generated PIL image.
        """
        pass

    @classmethod
    @abstractmethod
    def from_config(cls, config: ImageModelConfig) -> "BaseImageModel":
        """
        Constructs a model instance from the given configuration.

        Parameters
        ----------
        config : ImageModelConfig
            Model configuration object.

        Returns
        -------
        BaseImageModel
            An instance of a concrete image model class.
        """
        pass