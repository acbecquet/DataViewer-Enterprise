# DataViewer Enterprise — custom Postgres image with pg_cron
#
# Base postgres:16 does not bundle pg_cron. We install it from the official
# PGDG Debian package so init.sql's CREATE EXTENSION pg_cron succeeds.
# Pin to postgres:16 explicitly — major-version upgrades are deliberate events.
FROM postgres:16

RUN apt-get update \
 && apt-get install -y --no-install-recommends postgresql-16-cron \
 && rm -rf /var/lib/apt/lists/*
