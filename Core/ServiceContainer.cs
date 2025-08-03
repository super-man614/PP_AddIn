using System;
using System.Collections.Generic;
using PowerPointAddIn.Services;

namespace PowerPointAddIn.Core
{
    /// <summary>
    /// Simple dependency injection container for the PowerPoint add-in
    /// </summary>
    public class ServiceContainer
    {
        private readonly Dictionary<Type, object> _services = new Dictionary<Type, object>();
        private readonly Dictionary<Type, Func<object>> _factories = new Dictionary<Type, Func<object>>();
        private static ServiceContainer _instance;
        private static readonly object _lock = new object();

        /// <summary>
        /// Gets the singleton instance of the service container
        /// </summary>
        public static ServiceContainer Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = new ServiceContainer();
                    }
                }
                return _instance;
            }
        }

        private ServiceContainer()
        {
            RegisterDefaultServices();
        }

        /// <summary>
        /// Registers a singleton service instance
        /// </summary>
        public void RegisterSingleton<TInterface, TImplementation>(TImplementation instance)
            where TImplementation : class, TInterface
        {
            _services[typeof(TInterface)] = instance;
        }

        /// <summary>
        /// Registers a factory for creating service instances
        /// </summary>
        public void RegisterFactory<TInterface>(Func<TInterface> factory)
        {
            _factories[typeof(TInterface)] = () => factory();
        }

        /// <summary>
        /// Registers a transient service (creates new instance each time)
        /// </summary>
        public void RegisterTransient<TInterface, TImplementation>()
            where TImplementation : class, TInterface, new()
        {
            _factories[typeof(TInterface)] = () => new TImplementation();
        }

        /// <summary>
        /// Gets a service instance
        /// </summary>
        public T GetService<T>()
        {
            return (T)GetService(typeof(T));
        }

        /// <summary>
        /// Gets a service instance by type
        /// </summary>
        public object GetService(Type serviceType)
        {
            // Check for singleton instance
            if (_services.TryGetValue(serviceType, out var instance))
            {
                return instance;
            }

            // Check for factory
            if (_factories.TryGetValue(serviceType, out var factory))
            {
                return factory();
            }

            throw new InvalidOperationException($"Service of type {serviceType.Name} is not registered.");
        }

        /// <summary>
        /// Checks if a service is registered
        /// </summary>
        public bool IsRegistered<T>()
        {
            return IsRegistered(typeof(T));
        }

        /// <summary>
        /// Checks if a service is registered by type
        /// </summary>
        public bool IsRegistered(Type serviceType)
        {
            return _services.ContainsKey(serviceType) || _factories.ContainsKey(serviceType);
        }

        /// <summary>
        /// Registers default services for the application
        /// </summary>
        private void RegisterDefaultServices()
        {
            // Register core services
            var errorHandler = new ErrorHandlerService();
            RegisterSingleton<IErrorHandlerService, ErrorHandlerService>(errorHandler);

            // Register PowerPoint service
            RegisterFactory<IPowerPointService>(() => new PowerPointService(GetService<IErrorHandlerService>()));

            // Register slide service
            RegisterFactory<ISlideService>(() => new SlideService(
                GetService<IPowerPointService>(),
                GetService<IErrorHandlerService>()));

            // Register shape service
            RegisterFactory<IShapeService>(() => new ShapeService(
                GetService<IPowerPointService>(),
                GetService<IErrorHandlerService>()));
        }

        /// <summary>
        /// Clears all registered services (useful for testing)
        /// </summary>
        public void Clear()
        {
            _services.Clear();
            _factories.Clear();
        }

        /// <summary>
        /// Resets the container to default services
        /// </summary>
        public void Reset()
        {
            Clear();
            RegisterDefaultServices();
        }

        /// <summary>
        /// Disposes of any disposable services
        /// </summary>
        public void Dispose()
        {
            foreach (var service in _services.Values)
            {
                if (service is IDisposable disposable)
                {
                    try
                    {
                        disposable.Dispose();
                    }
                    catch (Exception ex)
                    {
                        // Log disposal errors but don't throw
                        System.Diagnostics.Debug.WriteLine($"Error disposing service: {ex.Message}");
                    }
                }
            }
            _services.Clear();
            _factories.Clear();
        }
    }
} 