import { CSSProperties, useState } from "react";
import EmptyState from "@shared/presentation/components/EmptyState";
import { Button, Image } from "@heroui/react";
import { IconArrowLeft, IconArrowRight, IconX } from "@tabler/icons-react";
import { kDefaultImage } from "@utils/constants";
import { FullscreenSlider } from "./Fullscreen_Slider";

const containerStyle: CSSProperties = {
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  flexDirection: "row" as "row",
};

export const ImageSlider = ({
  onPressRemove,
  images,
}: {
  images: string[];
  onPressRemove?: ((e: any, index: number) => void) | undefined;
}) => {
  if (images.length == 0) {
    return <EmptyState message="No images to display" />;
  }

  // Estado para llevar el control de la imagen actual
  const [currentIndex, setCurrentIndex] = useState(0);

  // Función para ir a la siguiente imagen
  const nextImage = () => {
    setCurrentIndex((prevIndex) =>
      prevIndex === images.length - 1 ? 0 : prevIndex + 1,
    );
  };

  // Función para ir a la imagen anterior
  const prevImage = () => {
    setCurrentIndex((prevIndex) =>
      prevIndex === 0 ? images.length - 1 : prevIndex - 1,
    );
  };

  return (
    <div className="slider-container w-full" style={containerStyle}>
      {images.length > 1 && (
        <Button
          className="bg-transparent border-none"
          size="md"
          isIconOnly
          onPress={prevImage}
        >
          <IconArrowLeft size={18} />
        </Button>
      )}

      <div key={currentIndex} className="relative">
        <Image
          src={images[currentIndex]}
          alt={`Preview ${currentIndex + 1}`}
          className="w-full h-64 object-cover rounded-md"
          onClick={(e) => {
            e.preventDefault();

            FullscreenSlider([images[currentIndex]]);
          }}
          fallbackSrc={kDefaultImage}
        />
        {onPressRemove != null && (
          <Button
            isIconOnly
            radius="full"
            size="sm"
            variant="light"
            color="danger"
            className="absolute top-2 right-2 z-10"
            onPress={(e) => onPressRemove(e, currentIndex)}
          >
            <IconX size="16" />
          </Button>
        )}
      </div>

      {images.length > 1 && (
        <Button
          className="bg-transparent border-none"
          size="md"
          isIconOnly
          onPress={nextImage}
        >
          <IconArrowRight size={18} />
        </Button>
      )}
    </div>
  );
};
