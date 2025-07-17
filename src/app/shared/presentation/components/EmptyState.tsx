import React from "react";

interface EmptyStateProps {
  message?: string;
  icon?: React.ReactNode;
}

const EmptyState: React.FC<EmptyStateProps> = ({
  message = "No data found.",
  icon,
}) => {
  return (
    <div className="flex flex-col items-center justify-center p-4  rounded-md">
      {icon ? (
        <div className="mb-4">{icon}</div>
      ) : (
        <svg
          xmlns="http://www.w3.org/2000/svg"
          className="h-12 w-12 text-gray-400 mb-4"
          fill="none"
          viewBox="0 0 24 24"
          stroke="currentColor"
          strokeWidth="2"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
          />
        </svg>
      )}
      <p className="text-gray-600 text-center">{message}</p>
    </div>
  );
};

export default EmptyState;
