## Summary

<!-- Describe what this PR does and why. -->

## Related issue

Closes #

## Changes

<!-- List the main changes introduced. -->

- 

## Testing

<!-- Describe how you tested this. Include Nextcloud version, PHP version, deployment type. -->

## Checklist

- [ ] No client-specific, proprietary, or personal data in code or comments
- [ ] New admin endpoints are decorated with `@AdminRequired`
- [ ] Sensitive values go through `ConfigService` and are encrypted at rest
- [ ] No new Composer runtime dependencies (or discussed in the issue first)
- [ ] `CHANGELOG.md` updated under `## Unreleased`
- [ ] PHP syntax verified (`find lib appinfo -name '*.php' | xargs php -l`)
