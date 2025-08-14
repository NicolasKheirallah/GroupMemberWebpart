import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { GraphService } from '../services/GraphService';
import styles from './GroupMembers.module.scss';

export interface ProfileImageProps {
  userId: string;
  graphService: GraphService;
  fallbackInitials: string;
  alt?: string;
  className?: string;
  size?: number;
}

const ProfileImage: React.FC<ProfileImageProps> = ({
  userId,
  graphService,
  fallbackInitials,
  alt = 'User profile image',
  className = '',
  size = 40
}) => {
  const [imgSrc, setImgSrc] = useState<string | null>(null);
  const [imageLoadError, setImageLoadError] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  useEffect(() => {
    let isCancelled = false;

    const loadPhoto = async (): Promise<void> => {
      try {
        setIsLoading(true);
        setImageLoadError(false);

        const photoUrl = await graphService.getUserPhoto(userId);

        if (!isCancelled) {
          if (photoUrl) {
            setImgSrc(photoUrl);
          } else {
            setImageLoadError(true);
          }
          setIsLoading(false);
        }
      } catch {
        if (!isCancelled) {
          setImageLoadError(true);
          setIsLoading(false);
        }
      }
    };

    loadPhoto().catch(console.error);

    return () => {
      isCancelled = true;
      if (imgSrc && imgSrc.startsWith('blob:')) {
        URL.revokeObjectURL(imgSrc);
      }
    };
  }, [userId, graphService]);

  const dynamicStyles = useMemo(() => ({
    width: `${size}px`,
    height: `${size}px`
  }), [size]);

  if (isLoading) {
    return (
      <div
        className={`${styles.defaultCoin} ${className}`}
        style={{
          ...dynamicStyles,
          opacity: 0.6,
          animation: 'pulse 1.5s ease-in-out infinite alternate'
        }}
        role="img"
        aria-label={`Loading profile image for ${fallbackInitials}`}
      >
        {fallbackInitials}
      </div>
    );
  }

  if (imgSrc && !imageLoadError) {
    return (
      <img
        src={imgSrc}
        alt={alt}
        className={`${styles.profileImage} ${className}`}
        style={dynamicStyles}
        onError={() => {
          setImageLoadError(true);
          if (imgSrc.startsWith('blob:')) {
            URL.revokeObjectURL(imgSrc);
          }
        }}
        loading="lazy"
        role="img"
        aria-label={`Profile image for ${alt || fallbackInitials}`}
      />
    );
  }

  return (
    <div
      className={`${styles.defaultCoin} ${className}`}
      style={dynamicStyles}
      role="img"
      aria-label={`Default avatar showing initials ${fallbackInitials}`}
      title={`No profile image available for ${alt || fallbackInitials}`}
    >
      {fallbackInitials}
    </div>
  );
};

export default ProfileImage;