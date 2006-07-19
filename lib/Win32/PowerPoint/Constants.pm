package Win32::PowerPoint::Constants;

use strict;
use Carp;

our $VERSION = '0.02';

our $AUTOLOAD;

sub new {
  my $class = shift;
  bless {

# ppSlideLayout
    ppLayoutBlank => 12,
    ppLayoutText  => 2,
    ppLayoutTitle => 1,

# ppAutoSize
    ppAutoSizeNone           => 0,
    ppAutoSizeShapeToFitText => 1,
    ppAutoSizeMixed          => -2,

# ppSaveAsFileType
    ppSaveAsPresentation => 1,
    ppSaveAsShow         => 7,

# ppParagraphAlignment
    ppAlignLeft       => 1,
    ppAlignCenter     => 2,
    ppAlignRight      => 3,
    ppAlignJustitfy   => 4,
    ppAlignDistribute => 5,
    ppAlignmentMixed  => -2,

# ppMouseActivation
    ppMouseClick => 1,
    ppMouseOver  => 2,

# msoTextOrientation
    msoTextOrientationHorizontal => 1,

# msoTriState
    msoTrue  => -1,
    msoFalse => 0,

  }, $class;
}

sub AUTOLOAD {
  my $self = shift;
  my $name = $AUTOLOAD;
  $name =~ s/.*://;
  if (exists $self->{$name})      { return $self->{$name}; }
  if (exists $self->{"pp$name"})  { return $self->{"pp$name"}; }
  if (exists $self->{"mso$name"}) { return $self->{"mso$name"}; }
  croak "constant $name does not exist";
}

sub DESTROY {}

1;
__END__

=head1 NAME

Win32::PowerPoint::Constants - Constants holder

=head1 DESCRIPTION

This is used internally in L<Win32::PowerPoint>.

=head1 METHOD

=head2 new

Creates an object.

=head1 SEE ALSO

PowerPoint's object browser and MSDN documentation.

=head1 AUTHOR

Kenichi Ishigaki, E<lt>ishigaki@cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2006 by Kenichi Ishigaki

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
