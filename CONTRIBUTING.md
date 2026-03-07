# Contributing to Outlook MCP Server

Thank you for your interest in contributing! This document provides guidelines for contributing to the project.

## Table of Contents
- [Getting Started](#getting-started)
- [Development Workflow](#development-workflow)
- [Code Standards](#code-standards)
- [Testing](#testing)
- [Pull Request Process](#pull-request-process)
- [Issue Reporting](#issue-reporting)

## Getting Started

### Prerequisites
- **Node.js** 18.0.0 or higher
- **macOS** (for AppleScript backend testing)
- **Microsoft Outlook** (classic or new)
- **Git**

### Setup Development Environment

1. **Fork and Clone**
   ```bash
   git clone https://github.com/YOUR_USERNAME/mcp-office365-mac.git
   cd mcp-office365-mac
   ```

2. **Install Dependencies**
   ```bash
   npm install
   ```

3. **Build**
   ```bash
   npm run build
   ```

4. **Run Tests**
   ```bash
   npm test
   ```

## Development Workflow

### Branch Naming
- `feature/description` - New features
- `fix/description` - Bug fixes
- `docs/description` - Documentation changes
- `refactor/description` - Code refactoring

### Local Testing with MCP Client

Test your changes with Claude Desktop or Claude Code:

```json
{
  "mcpServers": {
    "outlook-mac-dev": {
      "command": "node",
      "args": ["/path/to/mcp-office365-mac/dist/index.js"],
      "env": {
        "USE_GRAPH_API": "1"
      }
    }
  }
}
```

## Code Standards

This project uses strict TypeScript and ESLint configurations.

### TypeScript Rules
- **Strict mode enabled** - `strict: true`
- **No `any` types** - Use proper types or `unknown`
- **Explicit function return types** - All functions must declare return types
- **Strict boolean expressions** - No implicit boolean coercion
- **Prefer readonly** - Use `readonly` where applicable

### ESLint Rules
- No `console.log()` - Use `console.error()` or `console.warn()` only
- No explicit `any` types
- No floating promises - Always handle async operations
- Strict boolean expressions

### Run Linting
```bash
npm run lint        # Check for issues
npm run lint:fix    # Auto-fix issues
```

### Type Checking
```bash
npm run typecheck
```

## Testing

### Test Structure
- **Unit tests:** `tests/unit/` - Individual function/class testing
- **Integration tests:** `tests/integration/` - Multi-component testing
- **E2E tests:** `tests/e2e/` - Full MCP client testing

### Coverage Requirements
Minimum 80% coverage for:
- Lines
- Functions
- Branches
- Statements

### Test Commands
```bash
npm test                # Run all tests
npm run test:watch      # Watch mode
npm run test:coverage   # Generate coverage report
npm run test:ui         # Visual test UI
```

### Writing Tests
- Use Vitest for all tests
- Mock external dependencies (AppleScript, Microsoft Graph)
- Test both success and error cases
- Include edge cases

## Backend Testing Considerations

### AppleScript Backend
- Requires macOS
- Requires Outlook running
- Requires automation permissions: System Settings > Privacy & Security > Automation
- May require user interaction for first run

### Graph API Backend
- Requires Azure AD app registration (or use environment variable override)
- First run triggers device code authentication
- Subsequent runs use cached tokens

## Pull Request Process

1. **Create Feature Branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Changes**
   - Write code following standards
   - Add/update tests
   - Update documentation if needed

3. **Test Thoroughly**
   ```bash
   npm run lint
   npm run typecheck
   npm test
   npm run build
   ```

4. **Commit Changes**
   ```bash
   git add .
   git commit -m "feat: add your feature description"
   ```

   Use conventional commits:
   - `feat:` - New feature
   - `fix:` - Bug fix
   - `docs:` - Documentation
   - `refactor:` - Code refactoring
   - `test:` - Test updates
   - `chore:` - Build/tooling changes

5. **Push and Create PR**
   ```bash
   git push origin feature/your-feature-name
   ```
   - Open PR on GitHub
   - Fill out PR template
   - Link related issues
   - Request review

## Issue Reporting

### Before Creating an Issue
- Search existing issues to avoid duplicates
- Check troubleshooting section in README

### Bug Reports
Use the bug report template and include:
- **Environment:** macOS version, Outlook version, Node.js version
- **Backend:** AppleScript or Graph API
- **Steps to reproduce**
- **Expected vs actual behavior**
- **Error messages** (full stack trace if available)
- **MCP client:** Claude Desktop, Claude Code, or other

### Feature Requests
Use the feature request template and include:
- **Use case:** Why is this feature needed?
- **Proposed solution:** How should it work?
- **Alternatives considered:** Other approaches
- **Backend applicability:** AppleScript, Graph API, or both

## Support Policy

**This is a part-time, hobby project maintained by JBC Tech Solutions.**

### Response Times
- Bug reports: Best effort, typically 1-2 weeks
- Feature requests: Reviewed periodically
- Pull requests: Reviewed within 2-4 weeks
- Security issues: Prioritized, typically within 1 week

### Enterprise Support
If you need guaranteed response times or dedicated support:
- Consider creating your own Azure AD app registration
- Fork the repository for custom modifications
- Contact support@jbc.dev for consulting inquiries

## Community Resources
- **GitHub Discussions:** Ask questions, share ideas
- **GitHub Issues:** Bug reports and feature requests
- **README:** Installation and usage documentation
- **Code of Conduct:** [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)

## References
- [Model Context Protocol Documentation](https://modelcontextprotocol.io)
- [Microsoft Graph API Reference](https://learn.microsoft.com/en-us/graph/api/overview)
- [AppleScript Language Guide](https://developer.apple.com/library/archive/documentation/AppleScript/Conceptual/AppleScriptLangGuide/introduction/ASLR_intro.html)

## License
By contributing, you agree that your contributions will be licensed under the MIT License.

---

Thank you for contributing to Outlook MCP Server! 🎉
