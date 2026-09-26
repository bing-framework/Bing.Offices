# syntax=docker/dockerfile:1

# Keep the SDK image and the .NET 6 runtime installation reproducible. The
# runtime image currently used by this repository resolves to this digest.
FROM mcr.microsoft.com/dotnet/sdk@sha256:78235e09001f52b6592c458ac010775ebac6725422e80cd0c1650590f67b2743 AS runtime-base

ARG DOTNET_RUNTIME6_VERSION=6.0.36
ARG DOTNET_RUNTIME6_SHA512=afb6018fcabec468ccd7ae2f1131d8c9de7f4de7645b8f0c223efbbdbfdc515fb0642a399ebfe372c02044416c4cae463c9c802cd156b9da4181efff0e33ee94

ENV DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=true \
    DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 \
    NUGET_XMLDOC_MODE=skip

RUN set -eux; \
    runtime_archive="/tmp/dotnet-runtime-${DOTNET_RUNTIME6_VERSION}.tar.gz"; \
    curl --fail --silent --show-error --location \
      "https://builds.dotnet.microsoft.com/dotnet/Runtime/${DOTNET_RUNTIME6_VERSION}/dotnet-runtime-${DOTNET_RUNTIME6_VERSION}-linux-x64.tar.gz" \
      --output "$runtime_archive"; \
    printf '%s  %s\n' "$DOTNET_RUNTIME6_SHA512" "$runtime_archive" | sha512sum --check --status; \
    tar -xzf "$runtime_archive" --directory /usr/share/dotnet; \
    rm -f "$runtime_archive"; \
    useradd --create-home --shell /usr/sbin/nologin --uid 10001 bingoffices

WORKDIR /workspace

# This target intentionally has no font packages. It is used only for the
# negative font-environment check; rendering providers must not infer that a
# missing-font container is a valid rendering environment.
FROM runtime-base AS fontless
USER root
RUN set -eux; \
    for font_dir in /usr/share/fonts /usr/local/share/fonts; do \
      if [ -d "$font_dir" ]; then find "$font_dir" -type f -delete; fi; \
    done; \
    mkdir -p /usr/share/fonts /usr/local/share/fonts
USER bingoffices
ENTRYPOINT ["dotnet", "--info"]

FROM runtime-base AS verify
USER root

# Freeze the Debian package source and disable snapshot expiry checks. Package
# versions below are the versions present in the recorded snapshot.
RUN set -eux; \
    printf '%s\n' \
      'Types: deb' \
      'URIs: https://snapshot.debian.org/archive/debian/20260918T000000Z' \
      'Suites: bookworm' \
      'Components: main' \
      'Signed-By: /usr/share/keyrings/debian-archive-keyring.gpg' \
      '' \
      'Types: deb' \
      'URIs: https://snapshot.debian.org/archive/debian-security/20260918T000000Z' \
      'Suites: bookworm-security' \
      'Components: main' \
      'Signed-By: /usr/share/keyrings/debian-archive-keyring.gpg' \
      > /etc/apt/sources.list.d/debian.sources; \
    printf '%s\n' 'Acquire::Check-Valid-Until "false";' \
      > /etc/apt/apt.conf.d/99snapshot; \
    apt-get update; \
    apt-get install --yes --no-install-recommends \
      fontconfig=2.14.1-4 \
      fonts-dejavu-core=2.37-6 \
      fonts-noto-cjk=1:20220127+repack1-1; \
    fc-cache --force; \
    rm -rf /var/lib/apt/lists/*

USER bingoffices
ENTRYPOINT ["dotnet", "--info"]
