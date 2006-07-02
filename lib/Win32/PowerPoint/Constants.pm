package Win32::PowerPoint::Constants;

use strict;
use Carp;

our $VERSION = '0.01';

our $AUTOLOAD;

sub new {
  my $class = shift;
  bless {

# ppSlideLayout
    LayoutBlank => 12,
    LayoutText  => 2,
    LayoutTitle => 1,

# ppAutoSize
    AutoSizeNone           => 0,
    AutoSizeShapeToFitText => 1,
    AutoSizeMixed          => -2,

# ppSaveAsFileType
    SaveAsPresentation => 1,
    SaveAsShow         => 7,

# ppParagraphAlignment
    AlignLeft       => 1,
    AlignCenter     => 2,
    AlignRight      => 3,
    AlignJustitfy   => 4,
    AlignDistribute => 5,
    AlignmentMixed  => -2,

# msoTextOrientation
    TextOrientationHorizontal => 1,

# msoTriState
    True  => -1,
    False => 0,

  }, $class;
}

sub AUTOLOAD {
  my $self = shift;
  my $name = $AUTOLOAD;
  $name =~ s/.*://;
  if (exists $self->{$name}) { return $self->{$name}; }
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
