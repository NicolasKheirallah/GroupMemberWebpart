import { WebPartContext } from '@microsoft/sp-webpart-base';
import { LoggingService, ILoggingService } from './LoggingService';
import { ConfigurationService, IConfigurationService } from './ConfigurationService';
import { RetryService, IRetryService } from './RetryService';
import { CacheService, ICacheService } from './CacheService';
import { GroupMemberService, IGroupMemberService } from './GroupMemberService';
import { ProfileService, IProfileService } from './ProfileService';
import { SitePermissionService, ISitePermissionService } from './SitePermissionService';
import { UnifiedGraphService, IUnifiedGraphService } from './UnifiedGraphService';

export type ServiceLifetime = 'singleton' | 'transient' | 'scoped';

export interface ServiceDescriptor<T = unknown> {
  token: string | symbol;
  factory: (container: ServiceContainer) => T;
  lifetime: ServiceLifetime;
  instance?: T;
}

export interface IServiceContainer {
  register<T>(token: string | symbol, factory: (container: ServiceContainer) => T, lifetime?: ServiceLifetime): void;
  registerSingleton<T>(token: string | symbol, factory: (container: ServiceContainer) => T): void;
  registerTransient<T>(token: string | symbol, factory: (container: ServiceContainer) => T): void;
  registerScoped<T>(token: string | symbol, factory: (container: ServiceContainer) => T): void;
  resolve<T>(token: string | symbol): T;
  resolveOptional<T>(token: string | symbol): T | undefined;
  isRegistered(token: string | symbol): boolean;
  createScope(): IServiceContainer;
  dispose(): void;
}

// Service tokens for type-safe dependency injection
export const SERVICE_TOKENS = {
  LOGGING_SERVICE: Symbol('LoggingService'),
  CONFIGURATION_SERVICE: Symbol('ConfigurationService'),
  RETRY_SERVICE: Symbol('RetryService'),
  CACHE_SERVICE: Symbol('CacheService'),
  GROUP_MEMBER_SERVICE: Symbol('GroupMemberService'),
  PROFILE_SERVICE: Symbol('ProfileService'),
  SITE_PERMISSION_SERVICE: Symbol('SitePermissionService'),
  UNIFIED_GRAPH_SERVICE: Symbol('UnifiedGraphService'),
  WEB_PART_CONTEXT: Symbol('WebPartContext')
} as const;

export class ServiceContainer implements IServiceContainer {
  private services = new Map<string | symbol, ServiceDescriptor>();
  private scopedInstances = new Map<string | symbol, unknown>();
  private isDisposed = false;
  private parent?: ServiceContainer;

  constructor(parent?: ServiceContainer) {
    this.parent = parent;
  }

  public register<T>(
    token: string | symbol, 
    factory: (container: ServiceContainer) => T, 
    lifetime: ServiceLifetime = 'singleton'
  ): void {
    if (this.isDisposed) {
      throw new Error('Cannot register services on a disposed container');
    }

    this.services.set(token, {
      token,
      factory,
      lifetime
    });
  }

  public registerSingleton<T>(token: string | symbol, factory: (container: ServiceContainer) => T): void {
    this.register(token, factory, 'singleton');
  }

  public registerTransient<T>(token: string | symbol, factory: (container: ServiceContainer) => T): void {
    this.register(token, factory, 'transient');
  }

  public registerScoped<T>(token: string | symbol, factory: (container: ServiceContainer) => T): void {
    this.register(token, factory, 'scoped');
  }

  public resolve<T>(token: string | symbol): T {
    const service = this.resolveOptional<T>(token);
    if (service === undefined) {
      throw new Error(`Service with token '${String(token)}' is not registered`);
    }
    return service;
  }

  public resolveOptional<T>(token: string | symbol): T | undefined {
    if (this.isDisposed) {
      throw new Error('Cannot resolve services from a disposed container');
    }

    const descriptor = this.findDescriptor(token);
    if (!descriptor) {
      return undefined;
    }

    switch (descriptor.lifetime) {
      case 'singleton':
        return this.resolveSingleton<T>(descriptor);
      case 'transient':
        return this.resolveTransient<T>(descriptor);
      case 'scoped':
        return this.resolveScoped<T>(descriptor);
      default:
        throw new Error(`Unknown service lifetime: ${descriptor.lifetime}`);
    }
  }

  public isRegistered(token: string | symbol): boolean {
    return this.findDescriptor(token) !== undefined;
  }

  public createScope(): IServiceContainer {
    return new ServiceContainer(this);
  }

  public dispose(): void {
    if (this.isDisposed) {
      return;
    }

    // Dispose scoped instances
    for (const instance of this.scopedInstances.values()) {
      if (instance && typeof (instance as { dispose?: () => void }).dispose === 'function') {
        try {
          (instance as { dispose: () => void }).dispose();
        } catch (error) {
          console.warn('Error disposing scoped service:', error);
        }
      }
    }

    // If this is the root container, dispose singletons
    if (!this.parent) {
      for (const descriptor of this.services.values()) {
        if (descriptor.lifetime === 'singleton' && descriptor.instance) {
          if (typeof (descriptor.instance as { dispose?: () => void }).dispose === 'function') {
            try {
              (descriptor.instance as { dispose: () => void }).dispose();
            } catch (error) {
              console.warn('Error disposing singleton service:', error);
            }
          }
        }
      }
    }

    this.scopedInstances.clear();
    this.services.clear();
    this.isDisposed = true;
  }

  private findDescriptor(token: string | symbol): ServiceDescriptor | undefined {
    let descriptor = this.services.get(token);
    if (!descriptor && this.parent) {
      descriptor = this.parent.findDescriptor(token);
    }
    return descriptor;
  }

  private resolveSingleton<T>(descriptor: ServiceDescriptor): T {
    // For singletons, use the root container's instance
    const rootContainer = this.getRootContainer();
    const rootDescriptor = rootContainer.services.get(descriptor.token);
    
    if (!rootDescriptor) {
      throw new Error(`Singleton service descriptor not found in root container: ${String(descriptor.token)}`);
    }

    if (!rootDescriptor.instance) {
      rootDescriptor.instance = descriptor.factory(this);
    }
    
    return rootDescriptor.instance as T;
  }

  private resolveTransient<T>(descriptor: ServiceDescriptor): T {
    return descriptor.factory(this) as T;
  }

  private resolveScoped<T>(descriptor: ServiceDescriptor): T {
    if (!this.scopedInstances.has(descriptor.token)) {
      this.scopedInstances.set(descriptor.token, descriptor.factory(this));
    }
    return this.scopedInstances.get(descriptor.token) as T;
  }

  private getRootContainer(): ServiceContainer {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    let current: ServiceContainer = this;
    while (current.parent) {
      current = current.parent;
    }
    return current;
  }
}

// Global service container instance
let globalContainer: ServiceContainer | undefined;

export function createServiceContainer(context: WebPartContext): ServiceContainer {
  const container = new ServiceContainer();

  // Register WebPart context
  container.registerSingleton(SERVICE_TOKENS.WEB_PART_CONTEXT, () => context);

  // Register core services
  container.registerSingleton(SERVICE_TOKENS.LOGGING_SERVICE, (c) => 
    LoggingService.getInstance(c.resolve<WebPartContext>(SERVICE_TOKENS.WEB_PART_CONTEXT))
  );

  container.registerSingleton(SERVICE_TOKENS.CONFIGURATION_SERVICE, (c) => 
    ConfigurationService.getInstance(c.resolve<WebPartContext>(SERVICE_TOKENS.WEB_PART_CONTEXT))
  );

  container.registerSingleton(SERVICE_TOKENS.RETRY_SERVICE, () => 
    RetryService.getInstance()
  );

  container.registerSingleton(SERVICE_TOKENS.CACHE_SERVICE, () => 
    CacheService.getInstance()
  );

  // Register specialized services
  container.registerScoped(SERVICE_TOKENS.GROUP_MEMBER_SERVICE, (c) => 
    new GroupMemberService(c.resolve<WebPartContext>(SERVICE_TOKENS.WEB_PART_CONTEXT))
  );

  container.registerScoped(SERVICE_TOKENS.PROFILE_SERVICE, (c) => 
    new ProfileService(c.resolve<WebPartContext>(SERVICE_TOKENS.WEB_PART_CONTEXT))
  );

  container.registerScoped(SERVICE_TOKENS.SITE_PERMISSION_SERVICE, (c) => 
    new SitePermissionService(c.resolve<WebPartContext>(SERVICE_TOKENS.WEB_PART_CONTEXT))
  );

  container.registerScoped(SERVICE_TOKENS.UNIFIED_GRAPH_SERVICE, (c) => 
    new UnifiedGraphService(c.resolve<WebPartContext>(SERVICE_TOKENS.WEB_PART_CONTEXT))
  );

  return container;
}

export function getGlobalContainer(): ServiceContainer {
  if (!globalContainer) {
    throw new Error('Global service container has not been initialized. Call setGlobalContainer first.');
  }
  return globalContainer;
}

export function setGlobalContainer(container: ServiceContainer): void {
  if (globalContainer) {
    globalContainer.dispose();
  }
  globalContainer = container;
}

export function disposeGlobalContainer(): void {
  if (globalContainer) {
    globalContainer.dispose();
    globalContainer = undefined;
  }
}

// Decorator for automatic dependency injection
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function Injectable<T extends new (...args: any[]) => object>(constructor: T): T {
  return class extends constructor {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    constructor(...args: any[]) {
      super(...args);
      
      // Auto-inject dependencies if the class has an 'inject' method
      const instance = this as { inject?: (container: ServiceContainer) => void };
      if (typeof instance.inject === 'function') {
        instance.inject(getGlobalContainer());
      }
    }
  } as T;
}

// Decorator for injecting specific services
export function Inject(token: string | symbol): (target: unknown, propertyKey: string) => void {
  return function (target: unknown, propertyKey: string): void {
    // Store metadata about injected properties
    const targetWithProps = target as { _injectedProperties?: Array<{ propertyKey: string; token: string | symbol }> };
    if (!targetWithProps._injectedProperties) {
      targetWithProps._injectedProperties = [];
    }
    targetWithProps._injectedProperties.push({ propertyKey, token });
  };
}

// Helper function to resolve dependencies for decorated classes
export function resolveDependencies(instance: Record<string, unknown>, container: ServiceContainer): void {
  const proto = Object.getPrototypeOf(instance) as { _injectedProperties?: Array<{ propertyKey: string; token: string | symbol }> };
  if (proto._injectedProperties) {
    for (const { propertyKey, token } of proto._injectedProperties) {
      instance[propertyKey] = container.resolve(token);
    }
  }
}

// Type-safe service resolver functions
export function useLoggingService(container?: ServiceContainer): ILoggingService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<ILoggingService>(SERVICE_TOKENS.LOGGING_SERVICE);
}

export function useConfigurationService(container?: ServiceContainer): IConfigurationService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<IConfigurationService>(SERVICE_TOKENS.CONFIGURATION_SERVICE);
}

export function useRetryService(container?: ServiceContainer): IRetryService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<IRetryService>(SERVICE_TOKENS.RETRY_SERVICE);
}

export function useCacheService(container?: ServiceContainer): ICacheService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<ICacheService>(SERVICE_TOKENS.CACHE_SERVICE);
}

export function useGroupMemberService(container?: ServiceContainer): IGroupMemberService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<IGroupMemberService>(SERVICE_TOKENS.GROUP_MEMBER_SERVICE);
}

export function useProfileService(container?: ServiceContainer): IProfileService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<IProfileService>(SERVICE_TOKENS.PROFILE_SERVICE);
}

export function useSitePermissionService(container?: ServiceContainer): ISitePermissionService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<ISitePermissionService>(SERVICE_TOKENS.SITE_PERMISSION_SERVICE);
}

export function useUnifiedGraphService(container?: ServiceContainer): IUnifiedGraphService {
  const serviceContainer = container || getGlobalContainer();
  return serviceContainer.resolve<IUnifiedGraphService>(SERVICE_TOKENS.UNIFIED_GRAPH_SERVICE);
}

export default ServiceContainer;