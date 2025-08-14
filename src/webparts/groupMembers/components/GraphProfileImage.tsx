import * as React from 'react';
import { useState, useEffect, useCallback, useMemo } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './GroupMembers.module.scss';

export interface GraphProfileImageProps {
  userId: string;
  context: WebPartContext;
  fallbackInitials: string;
  alt?: string;
  className?: string;
  size?: number;
  disableAzureAD?: boolean; // Option to disable Graph API photo fetch
}

const GraphProfileImage: React.FC<GraphProfileImageProps> = ({
  userId, 
  context, 
  fallbackInitials, 
  alt = 'User profile image', 
  className = '', 
  size = 40,
  disableAzureAD = false
}) => {
  const [imgSrc, setImgSrc] = useState<string | null>(null);
  const [imageLoadError, setImageLoadError] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  // Memoize the photo fetching logic
  const loadPhoto = useCallback(async () => {
    // If disabled or no user ID, immediately fall back
    if (disableAzureAD || !userId) {
      setImageLoadError(true);
      setIsLoading(false);
      return;
    }

    // Check cache first
    const cacheKey = `profilePhoto_${userId}`;
    const cachedPhoto = sessionStorage.getItem(cacheKey);
    if (cachedPhoto) {
      setImgSrc(cachedPhoto);
      setIsLoading(false);
      return;
    }

    const controller = new AbortController();
    let isCancelled = false;

    try {
      // Token provider setup
      const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
      const token = await tokenProvider.getToken("https://graph.microsoft.com");
      
      // Profile photo URL
      const url = `https://graph.microsoft.com/v1.0/users/${userId}/photo/$value`;
      
      // Fetch with timeout and error handling
      const timeoutId = setTimeout(() => {
        isCancelled = true;
        controller.abort();
      }, 8000); // 8-second timeout

      const response = await fetch(url, {
        headers: { 
          Authorization: `Bearer ${token}`,
          'Accept': 'image/*'
        },
        signal: controller.signal
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        if (response.status === 404) {
          // User has no profile photo, cache the result
          sessionStorage.setItem(`${cacheKey}_noPhoto`, 'true');
        }
        throw new Error(`Photo fetch failed: ${response.status}`);
      }

      const buffer = await response.arrayBuffer();
      const contentType = response.headers.get("content-type") || "image/jpeg";
      const blob = new Blob([buffer], { type: contentType });
      const objectUrl = URL.createObjectURL(blob);

      if (!isCancelled) {
        setImgSrc(objectUrl);
        setIsLoading(false);
        // Cache the blob URL (with cleanup timer)
        sessionStorage.setItem(cacheKey, objectUrl);
        // Set cleanup timer for object URL
        setTimeout(() => {
          URL.revokeObjectURL(objectUrl);
          sessionStorage.removeItem(cacheKey);
        }, 10 * 60 * 1000); // 10 minutes
      }
    } catch (error) {
      if (!isCancelled) {
        setImageLoadError(true);
        setIsLoading(false);
        console.warn(`Failed to load profile photo for user ${userId}:`, error);
      }
    }

    // Cleanup function
    return () => {
      isCancelled = true;
    };
  }, [userId, context, disableAzureAD]);

  // Trigger photo loading effect
  useEffect(() => {
    let cleanup: (() => void) | undefined;

    const performLoad = async (): Promise<void> => {
      try {
        cleanup = await loadPhoto();
      } catch {
        setImageLoadError(true);
        setIsLoading(false);
      }
    };

    performLoad().catch(console.error);

    // Cleanup function
    return () => {
      if (cleanup) {
        cleanup();
      }
      if (imgSrc && imgSrc.startsWith('blob:')) {
        URL.revokeObjectURL(imgSrc);
      }
    };
  }, [loadPhoto]);

  // Memoized dynamic styles
  const dynamicStyles = useMemo(() => ({
    width: `${size}px`,
    height: `${size}px`
  }), [size]);

  // If the image is still loading, return a placeholder
  if (isLoading) {
    return (
      <div 
        className={`${styles.defaultCoin} ${className}`} 
        style={{
          ...dynamicStyles,
          opacity: 0.6,
          cursor: 'wait',
          animation: 'pulse 1.5s ease-in-out infinite alternate'
        }}
        role="img"
        aria-label={`Loading profile image for ${fallbackInitials}`}
      >
        {fallbackInitials}
      </div>
    );
  }

  // If the image loaded successfully, show it
  if (imgSrc && !imageLoadError) {
    return (
      <img
        src={imgSrc}
        alt={alt}
        className={`${styles.profileImage} ${className}`}
        style={dynamicStyles}
        onError={() => {
          setImageLoadError(true);
          // Clean up the failed image URL
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

  // If the image failed to load or was not found, render a fallback coin with the initials
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

export default GraphProfileImage;