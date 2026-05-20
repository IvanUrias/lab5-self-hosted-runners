# Laboratorio 5 — Self-hosted runners y troubleshooting

## 📋 Descripción

Este repositorio contiene la práctica del Laboratorio 5, enfocado en:

- **Self-hosted runners**: Configuración y gestión de runners en infraestructura propia
- **Workflows híbridos**: Pipelines que combinan runners hosted y self-hosted
- **Troubleshooting**: Resolución de errores comunes en GitHub Actions
- **Logging detallado**: Diagnóstico avanzado con logs verbose
- **Seguridad**: Análisis de riesgos en runners self-hosted

## 🏗️ Estructura del Proyecto

```
LAB 5/
├── .github/
│   └── workflows/
│       ├── hybrid-pipeline.yml          # Parte 2: Pipeline híbrido
│       ├── troubleshooting-errors.yml   # Parte 3: Errores deliberados
│       ├── troubleshooting-fixed.yml    # Parte 3: Errores corregidos
│       └── logging-detallado.yml        # Parte 4: Logging verbose
├── scripts/
│   └── health-check.ps1                # Script de diagnóstico
├── docs/
│   └── screenshots/                    # Capturas de pantalla
└── README.md
```

## 🚀 Workflows Disponibles

| Workflow | Tipo Runner | Descripción |
|----------|-------------|-------------|
| Pipeline Híbrido | Hosted + Self-hosted | Build en hosted, deploy en self-hosted |
| Troubleshooting Errores | Hosted + Self-hosted | Errores deliberados para práctica |
| Troubleshooting Fixed | Hosted + Self-hosted | Versiones corregidas |
| Logging Detallado | Self-hosted | Diagnóstico con logs verbose |

## 👤 Autor

Daniel Rodríguez García - daniel.rodriguezgarcia@tajamar365.com
