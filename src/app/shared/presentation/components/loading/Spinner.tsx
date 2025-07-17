import LoadingAnimation from "./Loading_animation";

export default function CustomSpinner() {
  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-[9999] animate-fade-in">
      <div className="bg-transparent p-6">
        <div className="flex items-center justify-center">
          {/* Spinner */}
          <LoadingAnimation />
        </div>
      </div>
    </div>
  );
}
