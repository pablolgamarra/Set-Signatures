# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0] - 2024-12-02

### Added

- Sistema de logging con niveles (INFO, WARNING, ERROR, SUCCESS)
- Estadísticas detalladas de ejecución
- Soporte para múltiples usuarios simultáneos
- Validación automática de plantillas HTML
- Manejo robusto de errores por usuario
- Barras de progreso durante la ejecución
- Parámetro -LogPath para guardar logs en archivo
- Documentación completa en README.md

### Changed

- Mejorado el manejo de conexiones a servicios de Microsoft
- Optimizado el procesamiento de arrays para evitar problemas con objetos individuales
- Actualizada la documentación con ejemplos prácticos

### Fixed

- Problema con objetos individuales vs arrays al procesar un solo usuario
- Error al procesar BusinessPhones vacíos
- Manejo correcto de placeholders en plantillas HTML

## [1.3] - 2024-04-13

### Added

- Versión inicial funcional del script
- Soporte básico para configuración de firmas
- Integración con Microsoft Graph y Exchange Online
