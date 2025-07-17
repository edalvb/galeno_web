import { useState, useEffect } from "react";
import ReactDOM from "react-dom/client";
import Slider, { Settings } from "react-slick";
import Zoom from "react-medium-image-zoom";
import "react-medium-image-zoom/dist/styles.css";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { Button, Tooltip } from "@heroui/react";
import { IconClipboardCopy, IconDownload, IconX } from "@tabler/icons-react";
import toast from "react-hot-toast";

interface FullscreenSliderProps {
  images: string[];
  onClose: () => void;
}

const FullscreenSliderComponent = (data: FullscreenSliderProps) => {
  const { images, onClose } = data;
  const [currentIndex, setCurrentIndex] = useState(0);
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    // Trigger animation after mount
    setMounted(true);
  }, []);

  const sliderSettings: Settings = {
    dots: true,
    infinite: true,
    speed: 500,
    slidesToShow: 1,
    slidesToScroll: 1,
    adaptiveHeight: true,
    afterChange: (current: number) => setCurrentIndex(current),
    dotsClass: "slick-dots slick-thumb",
    customPaging: (i: number) => (
      <a>
        <img src={images[i]} alt={`Thumbnail ${i}`} />
      </a>
    ),
  };

  const handleDownload = () => {
    window.open(images[currentIndex], "_blank");
  };

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(images[currentIndex]);
      toast.success("URL copied to clipboard.");
    } catch (error) {
      toast.error("Error copying URL");
    }
  };

  return (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        width: "100%",
        height: "100%",
        background: "rgba(0, 0, 0, 0.9)",
        zIndex: 1000,
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        opacity: mounted ? 1 : 0,
        transition: "opacity 0.5s ease-in-out",
      }}
    >
      <Tooltip content="Close">
        <Button
          onPress={onClose}
          isIconOnly
          style={{
            position: "absolute",
            top: "20px",
            right: "20px",
            zIndex: 1100,
          }}
        >
          <IconX size={18} />
        </Button>
      </Tooltip>

      <Tooltip content="Download">
        <Button
          isIconOnly
          onPress={handleDownload}
          style={{
            position: "absolute",
            top: "70px",
            right: "20px",
            zIndex: 1100,
          }}
        >
          <IconDownload size={18} />
        </Button>
      </Tooltip>

      <Tooltip content="Copy URL">
        <Button
          isIconOnly
          onPress={handleCopy}
          style={{
            position: "absolute",
            top: "120px",
            right: "20px",
            zIndex: 1100,
          }}
        >
          <IconClipboardCopy size={18} />
        </Button>
      </Tooltip>
      <div className="slider-container">
        <Slider {...sliderSettings}>
          {images.map((src, index) => (
            <div
              key={index}
              style={{
                textAlign: "center",
                position: "relative",
                display: "inline-block",
              }}
            >
              <Zoom>
                <img
                  src={src}
                  alt={`Image ${index}`}
                  style={{
                    maxHeight: "90vh",
                    margin: "0 auto",
                    display: "block",
                  }}
                />
              </Zoom>
            </div>
          ))}
        </Slider>
      </div>
    </div>
  );
};

// Function to mount the slider on the DOM
export function FullscreenSlider(images: string[]) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = ReactDOM.createRoot(container);

  // Close function to unmount and remove the slider
  const close = () => {
    root.unmount();
    container.remove();
  };

  root.render(<FullscreenSliderComponent images={images} onClose={close} />);
}
