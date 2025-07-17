import {
  IconStar,
  IconStarFilled,
  IconStarHalfFilled,
} from "@tabler/icons-react";

const RatingStars = ({ rating }: { rating: number }) => {
  const fullStars = Math.floor(rating);

  const hasHalfStar = rating - fullStars >= 0.5;

  const emptyStars = 5 - fullStars - (hasHalfStar ? 1 : 0);

  return (
    <div className="flex items-center ml-3">
      {Array.from({ length: fullStars }).map((_, i) => (
        <IconStarFilled key={`star-filled-${i}`} size="16" color="#e5be01" />
      ))}
      {hasHalfStar && (
        <IconStarHalfFilled key="star-half" size="16" color="#e5be01" />
      )}
      {Array.from({ length: emptyStars }).map((_, i) => (
        <IconStar key={`star-empty-${i}`} size="16" color="#e5be01" />
      ))}
    </div>
  );
};

export default RatingStars;
